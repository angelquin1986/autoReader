VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FichamedFamilTrbjOcu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdEmpleado As Long
    AnosMatrimonio As Integer
    NroHijos As Integer
    RelacIntraFamil As String
    RelacIntraFamilInt As Integer
    ServicioBasico As Boolean
    PoseeAnimales As Boolean
    TipoVivienda As Integer
    IngresoEconFamil As Currency
    Anteriores As String
    Actual As String
    TiempoTrabajoActual As String
    HorasTrabajoActual As Integer
    RotacionTrabActual As Boolean
    GradoAdaptFunciones As String
    BandRiesgoFisico As Boolean
    BandRiesgobiologico As Boolean
    BandRiesgoQuimico As Boolean
    BandRiesgoMeca As Boolean
    BandRiesgoErgo As Boolean
    BandRiesgoPsico  As Boolean
    UsoEquipoProtec As Boolean
    UsoEquipoProtecDet As String
    TieneSeguriTrab As Boolean
    TieneSeguriTrabDet As String
    RelacionCompanero As Integer
    AntiguedadOcup As String
    AntiguedadTrab As String
    RotacionPuesto As String
    RotacionTurnos As String
    OcupaActuales As String
    ActivExtraActual As String
    mobjEmpleado As PCProvCli
    ServicioBasicoLuz As Boolean
    ServicioBasicoAgua As Boolean
    ServicioBasicoTelefono As Boolean
    ServicioBasicoAlcantarillado As Boolean
    TipoViviendaVivienda As String
    TipoViviendaMaterial As String
    AdaptacionFunciones As Integer
    
    BandRiesgoFisico1 As Boolean
    BandRiesgoFisico2 As Boolean
    BandRiesgoFisico3 As Boolean
    BandRiesgobiologico1 As Boolean
    BandRiesgobiologico2 As Boolean
    BandRiesgobiologico3 As Boolean
    BandRiesgoQuimico1 As Boolean
    BandRiesgoQuimico2 As Boolean
    BandRiesgoQuimico3 As Boolean
    BandRiesgoMeca1 As Boolean
    BandRiesgoMeca2 As Boolean
    BandRiesgoMeca3 As Boolean
    BandRiesgoErgo1 As Boolean
    BandRiesgoErgo2 As Boolean
    BandRiesgoErgo3 As Boolean
    BandRiesgoPsico1  As Boolean
    BandRiesgoPsico2  As Boolean
    BandRiesgoPsico3  As Boolean
    
    
End Type
Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Friend Property Set mobjEmpleado(ByVal obj As PCProvCli)
    Set mProps.mobjEmpleado = obj
End Property

Public Property Get mobjEmpleado() As PCProvCli
    Set mobjEmpleado = mProps.mobjEmpleado
End Property

'/*
Public Property Let AdaptacionFunciones(ByVal value As Integer)
    mProps.AdaptacionFunciones = value
    mbooModificado = True
End Property

Public Property Get AdaptacionFunciones() As Integer
    AdaptacionFunciones = mProps.AdaptacionFunciones
End Property
'*/
Public Property Let AnosMatrimonio(ByVal value As Integer)
    mProps.AnosMatrimonio = value
    mbooModificado = True
End Property

Public Property Get AnosMatrimonio() As Integer
    AnosMatrimonio = mProps.AnosMatrimonio
End Property

Public Property Let NroHijos(ByVal value As Integer)
    mProps.NroHijos = value
    mbooModificado = True
End Property

Public Property Get NroHijos() As Integer
    NroHijos = mProps.NroHijos
End Property

Public Property Let RelacIntraFamil(ByVal value As String)
    'mProps.RelacIntraFamil = value
    'mbooModificado = True
End Property

Public Property Get RelacIntraFamil() As String
    'RelacIntraFamil = mProps.RelacIntraFamil
End Property
Public Property Let RelacIntraFamilInt(ByVal value As Integer)
    mProps.RelacIntraFamilInt = value
    mbooModificado = True
End Property

Public Property Get RelacIntraFamilInt() As Integer
    RelacIntraFamilInt = mProps.RelacIntraFamilInt
End Property

Public Property Let ServicioBasico(ByVal value As Boolean)
    mProps.ServicioBasico = value
    mbooModificado = True
End Property

Public Property Get ServicioBasico() As Boolean
    ServicioBasico = mProps.ServicioBasico
End Property

Public Property Let PoseeAnimales(ByVal value As Boolean)
    mProps.PoseeAnimales = value
    mbooModificado = True
End Property

Public Property Get PoseeAnimales() As Boolean
    PoseeAnimales = mProps.PoseeAnimales
End Property


Public Property Let TipoVivienda(ByVal value As Integer)
    mProps.TipoVivienda = value
    mbooModificado = True
End Property

Public Property Get TipoVivienda() As Integer
    TipoVivienda = mProps.TipoVivienda
End Property

Public Property Let IngresoEconFamil(ByVal value As Currency)
    mProps.IngresoEconFamil = value
    mbooModificado = True
End Property

Public Property Get IngresoEconFamil() As Currency
    IngresoEconFamil = mProps.IngresoEconFamil
End Property

Public Property Let Anteriores(ByVal value As String)
    mProps.Anteriores = value
    mbooModificado = True
End Property

Public Property Get Anteriores() As String
    Anteriores = mProps.Anteriores
End Property

Public Property Let Actual(ByVal value As String)
    mProps.Actual = value
    mbooModificado = True
End Property

Public Property Get Actual() As String
    Actual = mProps.Actual
End Property

Public Property Let TiempoTrabajoActual(ByVal value As String)
    mProps.TiempoTrabajoActual = value
    mbooModificado = True
End Property

Public Property Get TiempoTrabajoActual() As String
    TiempoTrabajoActual = mProps.TiempoTrabajoActual
End Property

Public Property Let HorasTrabajoActual(ByVal value As Integer)
    mProps.HorasTrabajoActual = value
    mbooModificado = True
End Property

Public Property Get HorasTrabajoActual() As Integer
    HorasTrabajoActual = mProps.HorasTrabajoActual
End Property

Public Property Let RotacionTrabActual(ByVal value As Boolean)
    mProps.RotacionTrabActual = value
    mbooModificado = True
End Property

Public Property Get RotacionTrabActual() As Boolean
    RotacionTrabActual = mProps.RotacionTrabActual
End Property

Public Property Let GradoAdaptFunciones(ByVal value As String)
    mProps.GradoAdaptFunciones = value
    mbooModificado = True
End Property

Public Property Get GradoAdaptFunciones() As String
    GradoAdaptFunciones = mProps.GradoAdaptFunciones
End Property

Public Property Let BandRiesgoFisico(ByVal value As Boolean)
    mProps.BandRiesgoFisico = value
    mbooModificado = True
End Property

Public Property Get BandRiesgoFisico() As Boolean
    BandRiesgoFisico = mProps.BandRiesgoFisico
End Property

Public Property Let BandRiesgobiologico(ByVal value As Boolean)
    mProps.BandRiesgobiologico = value
    mbooModificado = True
End Property

Public Property Get BandRiesgobiologico() As Boolean
    BandRiesgobiologico = mProps.BandRiesgobiologico
End Property

Public Property Let BandRiesgoQuimico(ByVal value As Boolean)
    mProps.BandRiesgoQuimico = value
    mbooModificado = True
End Property

Public Property Get BandRiesgoQuimico() As Boolean
    BandRiesgoQuimico = mProps.BandRiesgoQuimico
End Property

Public Property Let BandRiesgoMeca(ByVal value As Boolean)
    mProps.BandRiesgoMeca = value
    mbooModificado = True
End Property

Public Property Get BandRiesgoMeca() As Boolean
    BandRiesgoMeca = mProps.BandRiesgoMeca
End Property

Public Property Let BandRiesgoErgo(ByVal value As Boolean)
    mProps.BandRiesgoErgo = value
    mbooModificado = True
End Property

Public Property Get BandRiesgoErgo() As Boolean
    BandRiesgoErgo = mProps.BandRiesgoErgo
End Property

Public Property Let BandRiesgoPsico(ByVal value As Boolean)
    mProps.BandRiesgoPsico = value
    mbooModificado = True
End Property

Public Property Get BandRiesgoPsico() As Boolean
    BandRiesgoPsico = mProps.BandRiesgoPsico
End Property
Public Property Let UsoEquipoProtec(ByVal value As Boolean)
    mProps.UsoEquipoProtec = value
    mbooModificado = True
End Property

Public Property Get UsoEquipoProtec() As Boolean
    UsoEquipoProtec = mProps.UsoEquipoProtec
End Property

Public Property Let UsoEquipoProtecDet(ByVal value As String)
    mProps.UsoEquipoProtecDet = value
    mbooModificado = True
End Property

Public Property Get UsoEquipoProtecDet() As String
    UsoEquipoProtecDet = mProps.UsoEquipoProtecDet
End Property

Public Property Let TieneSeguriTrab(ByVal value As Boolean)
    mProps.TieneSeguriTrab = value
    mbooModificado = True
End Property

Public Property Get TieneSeguriTrab() As Boolean
    TieneSeguriTrab = mProps.TieneSeguriTrab
End Property

Public Property Let TieneSeguriTrabDet(ByVal value As String)
    mProps.TieneSeguriTrabDet = value
    mbooModificado = True
End Property

Public Property Get TieneSeguriTrabDet() As String
    TieneSeguriTrabDet = mProps.TieneSeguriTrabDet
End Property

Public Property Let ServicioBasicoLuz(ByVal value As Boolean)
    mProps.ServicioBasicoLuz = value
    mbooModificado = True
End Property

Public Property Get ServicioBasicoLuz() As Boolean
    ServicioBasicoLuz = mProps.ServicioBasicoLuz
End Property

Public Property Let ServicioBasicoAgua(ByVal value As Boolean)
    mProps.ServicioBasicoAgua = value
    mbooModificado = True
End Property

Public Property Get ServicioBasicoAgua() As Boolean
    ServicioBasicoAgua = mProps.ServicioBasicoAgua
End Property

Public Property Let ServicioBasicoTelefono(ByVal value As Boolean)
    mProps.ServicioBasicoTelefono = value
    mbooModificado = True
End Property

Public Property Get ServicioBasicoTelefono() As Boolean
    ServicioBasicoTelefono = mProps.ServicioBasicoTelefono
End Property

Public Property Let ServicioBasicoAlcantarillado(ByVal value As Boolean)
    mProps.ServicioBasicoAlcantarillado = value
    mbooModificado = True
End Property

Public Property Get ServicioBasicoAlcantarillado() As Boolean
    ServicioBasicoAlcantarillado = mProps.ServicioBasicoAlcantarillado
End Property

Public Property Let TipoViviendaVivienda(ByVal value As String)
    mProps.TipoViviendaVivienda = value
    mbooModificado = True
End Property

Public Property Get TipoViviendaVivienda() As String
    TipoViviendaVivienda = mProps.TipoViviendaVivienda
End Property

Public Property Let TipoViviendaMaterial(ByVal value As String)
    mProps.TipoViviendaMaterial = value
    mbooModificado = True
End Property

Public Property Get TipoViviendaMaterial() As String
    TipoViviendaMaterial = mProps.TipoViviendaMaterial
End Property

Public Property Let RelacionCompanero(ByVal value As Integer)
    mProps.RelacionCompanero = value
    mbooModificado = True
End Property

Public Property Get RelacionCompanero() As Integer
    RelacionCompanero = mProps.RelacionCompanero
End Property

Public Property Let AntiguedadOcup(ByVal value As String)
    mProps.AntiguedadOcup = value
    mbooModificado = True
End Property

Public Property Get AntiguedadOcup() As String
    AntiguedadOcup = mProps.AntiguedadOcup
End Property
Public Property Let AntiguedadTrab(ByVal value As String)
    mProps.AntiguedadTrab = value
    mbooModificado = True
End Property

Public Property Get AntiguedadTrab() As String
    AntiguedadTrab = mProps.AntiguedadTrab
End Property

Public Property Let RotacionPuesto(ByVal value As String)
    mProps.RotacionPuesto = value
    mbooModificado = True
End Property

Public Property Get RotacionPuesto() As String
    RotacionPuesto = mProps.RotacionPuesto
End Property

Public Property Let RotacionTurnos(ByVal value As String)
    mProps.RotacionTurnos = value
    mbooModificado = True
End Property

Public Property Get RotacionTurnos() As String
    RotacionTurnos = mProps.RotacionTurnos
End Property

Public Property Let OcupaActuales(ByVal value As String)
    mProps.OcupaActuales = value
    mbooModificado = True
End Property

Public Property Get OcupaActuales() As String
    OcupaActuales = mProps.OcupaActuales
End Property

Public Property Let ActivExtraActual(ByVal value As String)
    mProps.ActivExtraActual = value
    mbooModificado = True
End Property

Public Property Get ActivExtraActual() As String
    ActivExtraActual = mProps.ActivExtraActual
End Property


Public Property Get EsNuevo() As Boolean
    EsNuevo = mbooEsNuevo
End Property

Public Property Get Modificado() As Boolean
    Modificado = mbooModificado
End Property

Private Sub Class_Initialize()
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.mobjEmpleado = Nothing
End Sub

Public Sub Recuperar(ByVal id As Long)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT FichamedFamilTrbjOcu.* "
    sql = sql & " From FichamedFamilTrbjOcu "
    sql = sql & "WHERE idEmpleado = " & id
    Set rs = mProps.mobjEmpleado.Empresa.OpenRecordset(sql)
    With rs
        Do While Not .EOF
            mProps.IdEmpleado = !IdEmpleado
            If Not IsNull(!AnosMatrimonio) Then mProps.AnosMatrimonio = !AnosMatrimonio
            If Not IsNull(!NroHijos) Then mProps.NroHijos = !NroHijos
            If Not IsNull(!RelacIntraFamil) Then mProps.RelacIntraFamilInt = !RelacIntraFamil
            mProps.ServicioBasico = !ServicioBasico
            If Not IsNull(!PoseeAnimales) Then mProps.PoseeAnimales = !PoseeAnimales
            mProps.TipoVivienda = !TipoVivienda
            mProps.IngresoEconFamil = !IngresoEconFamil
            If Not IsNull(!Anteriores) Then mProps.Anteriores = !Anteriores
            If Not IsNull(!Actual) Then mProps.Actual = !Actual
            If Not IsNull(!TiempoTrabajoActual) Then mProps.TiempoTrabajoActual = !TiempoTrabajoActual
            If Not IsNull(!HorasTrabajoActual) Then mProps.HorasTrabajoActual = !HorasTrabajoActual
            If Not IsNull(!RotacionTrabActual) Then mProps.RotacionTrabActual = !RotacionTrabActual
            If Not IsNull(!GradoAdaptFunciones) Then mProps.GradoAdaptFunciones = !GradoAdaptFunciones
             mProps.BandRiesgoFisico = !BandRiesgoFisico
             mProps.BandRiesgoQuimico = !BandRiesgoQuimico
             mProps.BandRiesgobiologico = !BandRiesgobiologico
             mProps.BandRiesgoErgo = !BandRiesgoErgo
             mProps.BandRiesgoPsico = !BandRiesgoPsico
             mProps.BandRiesgoMeca = !BandRiesgoMeca
             mProps.UsoEquipoProtec = !UsoEquipoProtec
            If Not IsNull(!UsoEquipoProtecDet) Then mProps.UsoEquipoProtecDet = !UsoEquipoProtecDet
            mProps.TieneSeguriTrab = !TieneSeguriTrab
            If Not IsNull(!TieneSeguriTrabDet) Then mProps.TieneSeguriTrabDet = !TieneSeguriTrabDet
            mProps.RelacionCompanero = !RelacionCompanero
            If Not IsNull(!AntiguedadOcup) Then mProps.AntiguedadOcup = !AntiguedadOcup
            If Not IsNull(!AntiguedadTrab) Then mProps.AntiguedadTrab = !AntiguedadTrab
            If Not IsNull(!RotacionPuesto) Then mProps.RotacionPuesto = !RotacionPuesto
            If Not IsNull(!RotacionTurnos) Then mProps.RotacionTurnos = !RotacionTurnos
            If Not IsNull(!OcupaActuales) Then mProps.OcupaActuales = !OcupaActuales
            If Not IsNull(!ActivExtraActual) Then mProps.ActivExtraActual = !ActivExtraActual
            If Not IsNull(!ServicioBasicoLuz) Then mProps.ServicioBasicoLuz = !ServicioBasicoLuz
            If Not IsNull(!ServicioBasicoAgua) Then mProps.ServicioBasicoAgua = !ServicioBasicoAgua
            If Not IsNull(!ServicioBasicoTelefono) Then mProps.ServicioBasicoTelefono = !ServicioBasicoTelefono
            If Not IsNull(!ServicioBasicoAlcantarillado) Then mProps.ServicioBasicoAlcantarillado = !ServicioBasicoAlcantarillado
            If Not IsNull(!TipoViviendaVivienda) Then mProps.TipoViviendaVivienda = !TipoViviendaVivienda
            If Not IsNull(!TipoViviendaMaterial) Then mProps.TipoViviendaMaterial = !TipoViviendaMaterial
            If Not IsNull(!AdaptacionFunciones) Then mProps.AdaptacionFunciones = !AdaptacionFunciones
            
             mProps.BandRiesgoFisico1 = !BandRiesgoFisico1
             mProps.BandRiesgoFisico2 = !BandRiesgoFisico2
             mProps.BandRiesgoFisico3 = !BandRiesgoFisico3
             mProps.BandRiesgoQuimico1 = !BandRiesgoQuimico1
             mProps.BandRiesgoQuimico2 = !BandRiesgoQuimico2
             mProps.BandRiesgoQuimico3 = !BandRiesgoQuimico3
             mProps.BandRiesgobiologico1 = !BandRiesgobiologico1
             mProps.BandRiesgobiologico2 = !BandRiesgobiologico2
             mProps.BandRiesgobiologico3 = !BandRiesgobiologico3
             mProps.BandRiesgoErgo1 = !BandRiesgoErgo1
             mProps.BandRiesgoErgo2 = !BandRiesgoErgo2
             mProps.BandRiesgoErgo3 = !BandRiesgoErgo3
             mProps.BandRiesgoPsico1 = !BandRiesgoPsico1
             mProps.BandRiesgoPsico2 = !BandRiesgoPsico2
             mProps.BandRiesgoPsico3 = !BandRiesgoPsico3
             mProps.BandRiesgoMeca1 = !BandRiesgoMeca1
             mProps.BandRiesgoMeca2 = !BandRiesgoMeca2
             mProps.BandRiesgoMeca3 = !BandRiesgoMeca3
            
            
            rs.MoveNext
        Loop
            .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Grabar(ByVal IdEmpleado As Long)
    Dim sql As String, rs As Recordset
    On Error GoTo Etiq
    
 
            sql = "SELECT * FROM FichamedFamilTrbjOcu WHERE idempleado = " & IdEmpleado
            Set rs = mProps.mobjEmpleado.Empresa.OpenRecordsetParaEdit(sql)
        
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            sql = "SELECT * FROM FichamedFamilTrbjOcu WHERE 1=0"
            Set rs = mProps.mobjEmpleado.Empresa.OpenRecordsetParaEdit(sql)
            rs.AddNew
        Else
        End If
        
    With rs
            !IdEmpleado = IdEmpleado
            !AnosMatrimonio = mProps.AnosMatrimonio
             !NroHijos = mProps.NroHijos
             
             !RelacIntraFamil = mProps.RelacIntraFamilInt
            !ServicioBasico = mProps.ServicioBasico
             !PoseeAnimales = mProps.PoseeAnimales
            !TipoVivienda = mProps.TipoVivienda
            !IngresoEconFamil = mProps.IngresoEconFamil
            !Anteriores = mProps.Anteriores
            !Actual = mProps.Actual
            !TiempoTrabajoActual = mProps.TiempoTrabajoActual
            !HorasTrabajoActual = mProps.HorasTrabajoActual
            !RotacionTrabActual = mProps.RotacionTrabActual
            !GradoAdaptFunciones = mProps.GradoAdaptFunciones
             !BandRiesgoFisico = mProps.BandRiesgoFisico
             !BandRiesgoQuimico = mProps.BandRiesgoQuimico
             !BandRiesgobiologico = mProps.BandRiesgobiologico
             !BandRiesgoErgo = mProps.BandRiesgoErgo
             !BandRiesgoPsico = mProps.BandRiesgoPsico
             !BandRiesgoMeca = mProps.BandRiesgoMeca
             !UsoEquipoProtec = mProps.UsoEquipoProtec
             !UsoEquipoProtecDet = mProps.UsoEquipoProtecDet
            !TieneSeguriTrab = mProps.TieneSeguriTrab
             !TieneSeguriTrabDet = mProps.TieneSeguriTrabDet
            !RelacionCompanero = mProps.RelacionCompanero
            !AntiguedadOcup = mProps.AntiguedadOcup
            !AntiguedadTrab = mProps.AntiguedadTrab
            !RotacionPuesto = mProps.RotacionPuesto
            !RotacionTurnos = mProps.RotacionTurnos
            !OcupaActuales = mProps.OcupaActuales
            !ActivExtraActual = mProps.ActivExtraActual
            !ServicioBasicoLuz = mProps.ServicioBasicoLuz
            !ServicioBasicoAgua = mProps.ServicioBasicoAgua
            !ServicioBasicoTelefono = mProps.ServicioBasicoTelefono
            !ServicioBasicoAlcantarillado = mProps.ServicioBasicoAlcantarillado
            !TipoViviendaVivienda = mProps.TipoViviendaVivienda
            !TipoViviendaMaterial = mProps.TipoViviendaMaterial
            !AdaptacionFunciones = mProps.AdaptacionFunciones
            
             !BandRiesgoFisico1 = mProps.BandRiesgoFisico1
             !BandRiesgoQuimico1 = mProps.BandRiesgoQuimico1
             !BandRiesgobiologico1 = mProps.BandRiesgobiologico1
             !BandRiesgoErgo1 = mProps.BandRiesgoErgo1
             !BandRiesgoPsico1 = mProps.BandRiesgoPsico1
             !BandRiesgoMeca1 = mProps.BandRiesgoMeca1
            
             !BandRiesgoFisico2 = mProps.BandRiesgoFisico2
             !BandRiesgoQuimico2 = mProps.BandRiesgoQuimico2
             !BandRiesgobiologico2 = mProps.BandRiesgobiologico2
             !BandRiesgoErgo2 = mProps.BandRiesgoErgo2
             !BandRiesgoPsico2 = mProps.BandRiesgoPsico2
             !BandRiesgoMeca2 = mProps.BandRiesgoMeca2
            
             !BandRiesgoFisico3 = mProps.BandRiesgoFisico3
             !BandRiesgoQuimico3 = mProps.BandRiesgoQuimico3
             !BandRiesgobiologico3 = mProps.BandRiesgobiologico3
             !BandRiesgoErgo3 = mProps.BandRiesgoErgo3
             !BandRiesgoPsico3 = mProps.BandRiesgoPsico3
             !BandRiesgoMeca3 = mProps.BandRiesgoMeca3
            
            .Update
        .Close
    End With
    'GrabarHistorial (IdEmpleado)
    'GrabarCargasFam (IdEmpleado)
    'GrabarEstudios (IdEmpleado)
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    'Me.Empresa.RaiseRefreshEvent REFRESH_PROVCLI
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Public Sub Eliminar()
'    Dim sql As String, rs As Recordset
'
'        sql = "DELETE FROM fichamedicaemp WHERE idEmpleado = " & Me.IdEmpleado
'        'mProps.mobjEmpresa.
'    mbooEsNuevo = True
'    mbooModificado = True
    
End Sub
'


Public Property Let BandRiesgoFisico1(ByVal value As Boolean)
    mProps.BandRiesgoFisico1 = value
    mbooModificado = True
End Property

Public Property Get BandRiesgoFisico1() As Boolean
    BandRiesgoFisico1 = mProps.BandRiesgoFisico1
End Property

Public Property Let BandRiesgoFisico2(ByVal value As Boolean)
    mProps.BandRiesgoFisico2 = value
    mbooModificado = True
End Property

Public Property Get BandRiesgoFisico2() As Boolean
    BandRiesgoFisico2 = mProps.BandRiesgoFisico2
End Property

Public Property Let BandRiesgoFisico3(ByVal value As Boolean)
    mProps.BandRiesgoFisico3 = value
    mbooModificado = True
End Property

Public Property Get BandRiesgoFisico3() As Boolean
    BandRiesgoFisico3 = mProps.BandRiesgoFisico3
End Property

Public Property Let BandRiesgobiologico1(ByVal value As Boolean)
    mProps.BandRiesgobiologico1 = value
    mbooModificado = True
End Property

Public Property Get BandRiesgobiologico1() As Boolean
    BandRiesgobiologico1 = mProps.BandRiesgobiologico1
End Property

Public Property Let BandRiesgobiologico2(ByVal value As Boolean)
    mProps.BandRiesgobiologico2 = value
    mbooModificado = True
End Property

Public Property Get BandRiesgobiologico2() As Boolean
    BandRiesgobiologico2 = mProps.BandRiesgobiologico2
End Property

Public Property Let BandRiesgobiologico3(ByVal value As Boolean)
    mProps.BandRiesgobiologico3 = value
    mbooModificado = True
End Property

Public Property Get BandRiesgobiologico3() As Boolean
    BandRiesgobiologico3 = mProps.BandRiesgobiologico3
End Property

Public Property Let BandRiesgoQuimico1(ByVal value As Boolean)
    mProps.BandRiesgoQuimico1 = value
    mbooModificado = True
End Property

Public Property Get BandRiesgoQuimico1() As Boolean
    BandRiesgoQuimico1 = mProps.BandRiesgoQuimico1
End Property

Public Property Let BandRiesgoQuimico2(ByVal value As Boolean)
    mProps.BandRiesgoQuimico2 = value
    mbooModificado = True
End Property

Public Property Get BandRiesgoQuimico2() As Boolean
    BandRiesgoQuimico2 = mProps.BandRiesgoQuimico2
End Property

Public Property Let BandRiesgoQuimico3(ByVal value As Boolean)
    mProps.BandRiesgoQuimico3 = value
    mbooModificado = True
End Property

Public Property Get BandRiesgoQuimico3() As Boolean
    BandRiesgoQuimico3 = mProps.BandRiesgoQuimico3
End Property

Public Property Let BandRiesgoMeca1(ByVal value As Boolean)
    mProps.BandRiesgoMeca1 = value
    mbooModificado = True
End Property

Public Property Get BandRiesgoMeca1() As Boolean
    BandRiesgoMeca1 = mProps.BandRiesgoMeca1
End Property

Public Property Let BandRiesgoMeca2(ByVal value As Boolean)
    mProps.BandRiesgoMeca2 = value
    mbooModificado = True
End Property

Public Property Get BandRiesgoMeca2() As Boolean
    BandRiesgoMeca2 = mProps.BandRiesgoMeca2
End Property

Public Property Let BandRiesgoMeca3(ByVal value As Boolean)
    mProps.BandRiesgoMeca3 = value
    mbooModificado = True
End Property

Public Property Get BandRiesgoMeca3() As Boolean
    BandRiesgoMeca3 = mProps.BandRiesgoMeca3
End Property

Public Property Let BandRiesgoErgo1(ByVal value As Boolean)
    mProps.BandRiesgoErgo1 = value
    mbooModificado = True
End Property

Public Property Get BandRiesgoErgo1() As Boolean
    BandRiesgoErgo1 = mProps.BandRiesgoErgo1
End Property

Public Property Let BandRiesgoErgo2(ByVal value As Boolean)
    mProps.BandRiesgoErgo2 = value
    mbooModificado = True
End Property

Public Property Get BandRiesgoErgo2() As Boolean
    BandRiesgoErgo2 = mProps.BandRiesgoErgo2
End Property

Public Property Let BandRiesgoErgo3(ByVal value As Boolean)
    mProps.BandRiesgoErgo3 = value
    mbooModificado = True
End Property

Public Property Get BandRiesgoErgo3() As Boolean
    BandRiesgoErgo3 = mProps.BandRiesgoErgo3
End Property

Public Property Get BandRiesgoPsico1() As Boolean
    BandRiesgoPsico1 = mProps.BandRiesgoPsico1
End Property

Public Property Let BandRiesgoPsico1(ByVal value As Boolean)
    mProps.BandRiesgoPsico1 = value
    mbooModificado = True
End Property

Public Property Get BandRiesgoPsico2() As Boolean
    BandRiesgoPsico2 = mProps.BandRiesgoPsico2
End Property

Public Property Let BandRiesgoPsico2(ByVal value As Boolean)
    mProps.BandRiesgoPsico2 = value
    mbooModificado = True
End Property

Public Property Get BandRiesgoPsico3() As Boolean
    BandRiesgoPsico3 = mProps.BandRiesgoPsico3
End Property

Public Property Let BandRiesgoPsico3(ByVal value As Boolean)
    mProps.BandRiesgoPsico3 = value
    mbooModificado = True
End Property

