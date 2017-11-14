VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FichaMedicaExamenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdExmFisico As Integer
    IdEmpleado As Integer
    IdNroHistorial As Integer
    FechaConsulta As Date
    HoraConsulta As Date
    MotivoConsulta As String
    TipoHistorial As Integer
    Peso As String
    Talla As String
    TA As String
    FC As String
    temp As String
    Biotipo As String
    StdNutricional As String
    AgudezaVisual As Boolean
    RejillaAmsler As Boolean
    UsaLentes As Boolean
    Receta As String
    PielFaneras As String
    Cabeza As String
    Ojos As String
    EnfermVista As String
    Oidos As String
    OtoscopiaOD As String
    OtoscopiaOI As String
    AudiometriaOD As String
    AudiometriaOI As String
    Nariz As String
    Boca As String
    Dentadura As String
    Cuello As String
    Torax As String
    CsPs As String
    Garganta As String
    Espirometria As String
    RsCs As String
    EKG As String
    PrubEsfuerzo As String
    Abdomen As String
    AparatoUrogenital As String
    RegionAno As String '-----------
    ColumnaVert As String
    EvaPostural As String
    ExtSuperior As String
    ExtInferiores As String
    Motilidad As String
    Sensibilidad As String
    Reflejos As String
    Coordinacion As String
    MovInvol As String
    Equilibrio As String
    Conciencia As String
    Memoria As String
    GangliosLinf As String
    Sangre As String
    Emo As String
    Heces As String
    RX As String
    TAC As String
    Otros As String
    Nota As String
    ImpDiagnostica As String
    AptoSinRes As String
    AptoConResPer As String
    AptoConResLabAdap As String
    AptoConResLabRes As String
    AptoConResObs As String
    AptoConResNoApto As String
    Prescripcion As String
    FechaConsVisual As Date
    FechaConsAudi As Date
    FechaPresBronco As Date
    FechaPresToxi As Date
    FechaPresInf As Date
    FechaExamen As Date
    FechaNacimiento As Date
    ' /* Consulta Diaria
    FechaConsultaDiaria As Date
    DiagnosticoConsultaDiaria As String
    PrescripcionConsultaDiaria As String
    
    ConsultaDiaria As Boolean
    ' */
    ' /* Fechas Tipo Historial
    FechaPreocupacional As Date
    FechaOcupacional As Date
    FechaReintegro As Date
    FechaRetiro As Date
    ' */
    '/* Agudeza y rejilla
    AgudezaVisualOD As String
    AgudezaVisualOI As String
    RejillaAmslerOD As String
    RejillaAmslerOI As String
    '*/
    ' */
    CheckAptoConResPer As Boolean
    CheckConResLabAdap As Boolean
    CheckAptoConResLabRes As Boolean
    CheckAptoConResObs As Boolean
    CheckAptoConResNoApto As Boolean
    Edad As Integer
    ' /*
    idEnfermedad As Long
    CodEnfermedad As String
    DescripcionEnfermedad As String
    ' */
    
    RecetaPrescrip As String
    
    mobjEmpleado As PCProvCli
    
End Type
Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean
Event Refresh(Tipo As Integer)
Private mcolEnfDetalle As Collection

' /* Cargar Enfermedades */
Public Property Let idEnfermedad(ByVal value As Long)
    mProps.idEnfermedad = value
    mbooModificado = True
End Property

Public Property Get idEnfermedad() As Long
    idEnfermedad = mProps.idEnfermedad
End Property

Public Property Let CodEnfermedad(ByVal value As String)
    mProps.CodEnfermedad = value
    mbooModificado = True
End Property

Public Property Get CodEnfermedad() As String
    CodEnfermedad = mProps.CodEnfermedad
End Property

Public Property Let DescripcionEnfermedad(ByVal value As String)
    mProps.DescripcionEnfermedad = value
    mbooModificado = True
End Property

Public Property Get DescripcionEnfermedad() As String
    DescripcionEnfermedad = mProps.DescripcionEnfermedad
End Property
' /* Cargar Enfermedades */

' // Agudeza y rejilla - INICIO
Public Property Let AgudezaVisualOD(ByVal value As String)
    mProps.AgudezaVisualOD = value
    mbooModificado = True
End Property

Public Property Get AgudezaVisualOD() As String
    AgudezaVisualOD = mProps.AgudezaVisualOD
End Property

Public Property Let AgudezaVisualOI(ByVal value As String)
    mProps.AgudezaVisualOI = value
    mbooModificado = True
End Property

Public Property Get AgudezaVisualOI() As String
    AgudezaVisualOI = mProps.AgudezaVisualOI
End Property

Public Property Let RejillaAmslerOD(ByVal value As String)
    mProps.RejillaAmslerOD = value
    mbooModificado = True
End Property

Public Property Get RejillaAmslerOD() As String
    RejillaAmslerOD = mProps.RejillaAmslerOD
End Property

Public Property Let RejillaAmslerOI(ByVal value As String)
    mProps.RejillaAmslerOI = value
    mbooModificado = True
End Property

Public Property Get RejillaAmslerOI() As String
    RejillaAmslerOI = mProps.RejillaAmslerOI
End Property
' // Agudeza y rejilla - FIN

' // Ficha Medica Consulta Diaria - INICIO
Public Property Let FechaConsultaDiaria(ByVal value As Date)
    mProps.FechaConsultaDiaria = value
    mbooModificado = True
End Property

Public Property Get FechaConsultaDiaria() As Date
    FechaConsultaDiaria = mProps.FechaConsultaDiaria
End Property

Public Property Let DiagnosticoConsultaDiaria(ByVal value As String)
    mProps.DiagnosticoConsultaDiaria = value
    mbooModificado = True
End Property

Public Property Get DiagnosticoConsultaDiaria() As String
    DiagnosticoConsultaDiaria = mProps.FechaConsultaDiaria
End Property

Public Property Let PrescripcionConsultaDiaria(ByVal value As String)
    mProps.PrescripcionConsultaDiaria = value
    mbooModificado = True
End Property

Public Property Get PrescripcionConsultaDiaria() As String
    PrescripcionConsultaDiaria = mProps.PrescripcionConsultaDiaria
End Property

Public Property Let ConsultaDiaria(ByVal value As Boolean)
    mProps.ConsultaDiaria = value
    mbooModificado = True
End Property

Public Property Get ConsultaDiaria() As Boolean
    ConsultaDiaria = mProps.ConsultaDiaria
End Property

' // Ficha Medica Consulta Diaria - FIN

' // Ficha Medica Historia Clinica Fechas - INICIO
Public Property Let FechaPreocupacional(ByVal value As Date)
    mProps.FechaPreocupacional = value
    mbooModificado = True
End Property

Public Property Get FechaPreocupacional() As Date
    FechaPreocupacional = mProps.FechaPreocupacional
End Property

Public Property Let FechaOcupacional(ByVal value As Date)
    mProps.FechaOcupacional = value
    mbooModificado = True
End Property

Public Property Get FechaOcupacional() As Date
    FechaOcupacional = mProps.FechaOcupacional
End Property

Public Property Let FechaReintegro(ByVal value As Date)
    mProps.FechaReintegro = value
    mbooModificado = True
End Property

Public Property Get FechaReintegro() As Date
    FechaReintegro = mProps.FechaReintegro
End Property

Public Property Let FechaRetiro(ByVal value As Date)
    mProps.FechaRetiro = value
    mbooModificado = True
End Property

Public Property Get FechaRetiro() As Date
    FechaRetiro = mProps.FechaRetiro
End Property

' // Ficha Medica Historia Clinica Fechas - FIN

Public Property Let IdEmpleado(ByVal value As Long)
    mProps.IdEmpleado = value
    mbooModificado = True
End Property

Public Property Get IdEmpleado() As Long
    IdEmpleado = mProps.IdEmpleado
End Property

Public Property Get IdExmFisico() As Long
    IdExmFisico = mProps.IdExmFisico
End Property

Friend Property Set mobjEmpleado(ByVal obj As PCProvCli)
    Set mProps.mobjEmpleado = obj
End Property

Public Property Get mobjEmpleado() As PCProvCli
    Set mobjEmpleado = mProps.mobjEmpleado
End Property

Public Property Let IdNroHistorial(ByVal value As Integer)
    mProps.IdNroHistorial = value
    mbooModificado = True
End Property

Public Property Get IdNroHistorial() As Integer
    IdNroHistorial = mProps.IdNroHistorial
End Property

Public Property Let HoraConsulta(ByVal value As Date)
    If Not IsDate(value) Then Err.Raise ERR_INVALIDO, "FichaMedicaExamenes.HoraConsulta ", MSGERR_INVALIDO
    mProps.HoraConsulta = value
    mbooModificado = True
End Property

Public Property Get HoraConsulta() As Date
    HoraConsulta = mProps.HoraConsulta
End Property

Public Property Let Peso(ByVal value As String)
    mProps.Peso = value
    mbooModificado = True
End Property

Public Property Get Peso() As String
    Peso = mProps.Peso
End Property

Public Property Let TipoHistorial(ByVal value As Integer)
    mProps.TipoHistorial = value
    mbooModificado = True
End Property

Public Property Get TipoHistorial() As Integer
    TipoHistorial = mProps.TipoHistorial
End Property

Public Property Let FechaConsulta(ByVal value As Date)
    If Not IsDate(value) Then Err.Raise ERR_INVALIDO, "FichaMedicaExamenes.FechaConsulta", MSGERR_INVALIDO
    mProps.FechaConsulta = value
    mbooModificado = True
End Property

Public Property Get FechaConsulta() As Date
    FechaConsulta = mProps.FechaConsulta
End Property

Public Property Let MotivoConsulta(ByVal value As String)
    mProps.MotivoConsulta = value
    mbooModificado = True
End Property

Public Property Get MotivoConsulta() As String
    MotivoConsulta = mProps.MotivoConsulta
End Property

Public Property Let Talla(ByVal value As String)
    mProps.Talla = value
    mbooModificado = True
End Property

Public Property Get Talla() As String
    Talla = mProps.Talla
End Property

Public Property Let TA(ByVal value As String)
    mProps.TA = value
    mbooModificado = True
End Property

Public Property Get TA() As String
    TA = mProps.TA
End Property

Public Property Let FC(ByVal value As String)
    mProps.FC = value
    mbooModificado = True
End Property

Public Property Get FC() As String
    FC = mProps.FC
End Property

Public Property Let temp(ByVal value As String)
    mProps.temp = value
    mbooModificado = True
End Property

Public Property Get temp() As String
    temp = mProps.temp
End Property

Public Property Let Biotipo(ByVal value As String)
    mProps.Biotipo = value
    mbooModificado = True
End Property

Public Property Get Biotipo() As String
    Biotipo = mProps.Biotipo
End Property

Public Property Let StdNutricional(ByVal value As String)
    mProps.StdNutricional = value
    mbooModificado = True
End Property

Public Property Get StdNutricional() As String
    StdNutricional = mProps.StdNutricional
End Property
Public Property Let AgudezaVisual(ByVal value As Boolean)
    mProps.AgudezaVisual = value
    mbooModificado = True
End Property

Public Property Get AgudezaVisual() As Boolean
    AgudezaVisual = mProps.AgudezaVisual
End Property

Public Property Let RejillaAmsler(ByVal value As Boolean)
    mProps.RejillaAmsler = value
    mbooModificado = True
End Property

Public Property Get RejillaAmsler() As Boolean
    RejillaAmsler = mProps.RejillaAmsler
End Property

Public Property Let UsaLentes(ByVal value As Boolean)
    mProps.UsaLentes = value
    mbooModificado = True
End Property

Public Property Get UsaLentes() As Boolean
    UsaLentes = mProps.UsaLentes
End Property

Public Property Let Prescripcion(ByVal value As String)
    mProps.Prescripcion = value
    mbooModificado = True
End Property

Public Property Get Prescripcion() As String
    Prescripcion = mProps.Prescripcion
End Property

Public Property Let Receta(ByVal value As String)
    mProps.Receta = value
    mbooModificado = True
End Property

Public Property Get Receta() As String
    Receta = mProps.Receta
End Property

Public Property Let PielFaneras(ByVal value As String)
    mProps.PielFaneras = value
    mbooModificado = True
End Property

Public Property Get PielFaneras() As String
    PielFaneras = mProps.PielFaneras
End Property

Public Property Let Cabeza(ByVal value As String)
    mProps.Cabeza = value
    mbooModificado = True
End Property

Public Property Get Cabeza() As String
    Cabeza = mProps.Cabeza
End Property

Public Property Let Ojos(ByVal value As String)
    mProps.Ojos = value
    mbooModificado = True
End Property

Public Property Get Ojos() As String
    Ojos = mProps.Ojos
End Property

Public Property Let EnfermVista(ByVal value As String)
    mProps.EnfermVista = value
    mbooModificado = True
End Property

Public Property Get EnfermVista() As String
    EnfermVista = mProps.EnfermVista
End Property

Public Property Let Oidos(ByVal value As String)
    mProps.Oidos = value
    mbooModificado = True
End Property

Public Property Get Oidos() As String
    Oidos = mProps.Oidos
End Property

Public Property Let OtoscopiaOD(ByVal value As String)
    mProps.OtoscopiaOD = value
    mbooModificado = True
End Property

Public Property Get OtoscopiaOD() As String
    OtoscopiaOD = mProps.OtoscopiaOD
End Property

Public Property Let OtoscopiaOI(ByVal value As String)
    mProps.OtoscopiaOI = value
    mbooModificado = True
End Property

Public Property Get OtoscopiaOI() As String
    OtoscopiaOI = mProps.OtoscopiaOI
End Property

Public Property Let AudiometriaOD(ByVal value As String)
    mProps.AudiometriaOD = value
    mbooModificado = True
End Property

Public Property Get AudiometriaOD() As String
    AudiometriaOD = mProps.AudiometriaOD
End Property

Public Property Let AudiometriaOI(ByVal value As String)
    mProps.AudiometriaOI = value
    mbooModificado = True
End Property

Public Property Get AudiometriaOI() As String
    AudiometriaOI = mProps.AudiometriaOI
End Property

Public Property Let Nariz(ByVal value As String)
    mProps.Nariz = value
    mbooModificado = True
End Property

Public Property Get Nariz() As String
    Nariz = mProps.Nariz
End Property

Public Property Let Boca(ByVal value As String)
    mProps.Boca = value
    mbooModificado = True
End Property

Public Property Get Boca() As String
    Boca = mProps.Boca
End Property

Public Property Let Dentadura(ByVal value As String)
    mProps.Dentadura = value
    mbooModificado = True
End Property

Public Property Get Dentadura() As String
    Dentadura = mProps.Dentadura
End Property

Public Property Let Cuello(ByVal value As String)
    mProps.Cuello = value
    mbooModificado = True
End Property

Public Property Get Cuello() As String
    Cuello = mProps.Cuello
End Property

Public Property Let Torax(ByVal value As String)
    mProps.Torax = value
    mbooModificado = True
End Property

Public Property Get Torax() As String
    Torax = mProps.Torax
End Property

Public Property Let CsPs(ByVal value As String)
    mProps.CsPs = value
    mbooModificado = True
End Property

Public Property Get CsPs() As String
    CsPs = mProps.CsPs
End Property

Public Property Let Garganta(ByVal value As String)
    mProps.Garganta = value
    mbooModificado = True
End Property

Public Property Get Garganta() As String
    Garganta = mProps.Garganta
End Property

Public Property Let Espirometria(ByVal value As String)
    mProps.Espirometria = value
    mbooModificado = True
End Property

Public Property Get Espirometria() As String
    Espirometria = mProps.Espirometria
End Property

Public Property Let RsCs(ByVal value As String)
    mProps.RsCs = value
    mbooModificado = True
End Property

Public Property Get RsCs() As String
    RsCs = mProps.RsCs
End Property

Public Property Let EKG(ByVal value As String)
    mProps.EKG = value
    mbooModificado = True
End Property

Public Property Get EKG() As String
    EKG = mProps.EKG
End Property

Public Property Let PrubEsfuerzo(ByVal value As String)
    mProps.PrubEsfuerzo = value
    mbooModificado = True
End Property

Public Property Get PrubEsfuerzo() As String
    PrubEsfuerzo = mProps.PrubEsfuerzo
End Property

Public Property Let Abdomen(ByVal value As String)
    mProps.Abdomen = value
    mbooModificado = True
End Property

Public Property Get Abdomen() As String
    Abdomen = mProps.Abdomen
End Property

Public Property Let AparatoUrogenital(ByVal value As String)
    mProps.AparatoUrogenital = value
    mbooModificado = True
End Property

Public Property Get AparatoUrogenital() As String
    AparatoUrogenital = mProps.AparatoUrogenital
End Property

Public Property Let RegionAno(ByVal value As String)
    mProps.RegionAno = value
    mbooModificado = True
End Property

Public Property Get RegionAno() As String
    RegionAno = mProps.RegionAno
End Property

Public Property Let ColumnaVert(ByVal value As String)
    mProps.ColumnaVert = value
    mbooModificado = True
End Property

Public Property Get ColumnaVert() As String
    ColumnaVert = mProps.ColumnaVert
End Property

Public Property Let EvaPostural(ByVal value As String)
    mProps.EvaPostural = value
    mbooModificado = True
End Property

Public Property Get EvaPostural() As String
    EvaPostural = mProps.EvaPostural
End Property
Public Property Let ExtSuperior(ByVal value As String)
    mProps.ExtSuperior = value
    mbooModificado = True
End Property

Public Property Get ExtSuperior() As String
    ExtSuperior = mProps.ExtSuperior
End Property

Public Property Let ExtInferiores(ByVal value As String)
    mProps.ExtInferiores = value
    mbooModificado = True
End Property

Public Property Get ExtInferiores() As String
    ExtInferiores = mProps.ExtInferiores
End Property

Public Property Let Motilidad(ByVal value As String)
    mProps.Motilidad = value
    mbooModificado = True
End Property

Public Property Get Motilidad() As String
    Motilidad = mProps.Motilidad
End Property

Public Property Let Sensibilidad(ByVal value As String)
    mProps.Sensibilidad = value
    mbooModificado = True
End Property

Public Property Get Sensibilidad() As String
    Sensibilidad = mProps.Sensibilidad
End Property

Public Property Let Reflejos(ByVal value As String)
    mProps.Reflejos = value
    mbooModificado = True
End Property

Public Property Get Reflejos() As String
    Reflejos = mProps.Reflejos
End Property

Public Property Let Coordinacion(ByVal value As String)
    mProps.Coordinacion = value
    mbooModificado = True
End Property

Public Property Get Coordinacion() As String
    Coordinacion = mProps.Coordinacion
End Property

Public Property Let MovInvol(ByVal value As String)
    mProps.MovInvol = value
    mbooModificado = True
End Property

Public Property Get MovInvol() As String
    MovInvol = mProps.MovInvol
End Property

Public Property Let Equilibrio(ByVal value As String)
    mProps.Equilibrio = value
    mbooModificado = True
End Property

Public Property Get Equilibrio() As String
    Equilibrio = mProps.Equilibrio
End Property

Public Property Let Conciencia(ByVal value As String)
    mProps.Conciencia = value
    mbooModificado = True
End Property

Public Property Get Conciencia() As String
    Conciencia = mProps.Conciencia
End Property

Public Property Let Memoria(ByVal value As String)
    mProps.Memoria = value
    mbooModificado = True
End Property

Public Property Get Memoria() As String
    Memoria = mProps.Memoria
End Property

Public Property Let GangliosLinf(ByVal value As String)
    mProps.GangliosLinf = value
    mbooModificado = True
End Property

Public Property Get GangliosLinf() As String
    GangliosLinf = mProps.GangliosLinf
End Property

Public Property Let Sangre(ByVal value As String)
    mProps.Sangre = value
    mbooModificado = True
End Property

Public Property Get Sangre() As String
    Sangre = mProps.Sangre
End Property

Public Property Let Emo(ByVal value As String)
    mProps.Emo = value
    mbooModificado = True
End Property

Public Property Get Emo() As String
    Emo = mProps.Emo
End Property

Public Property Let Heces(ByVal value As String)
    mProps.Heces = value
    mbooModificado = True
End Property

Public Property Get Heces() As String
    Heces = mProps.Heces
End Property

Public Property Let RX(ByVal value As String)
    mProps.RX = value
    mbooModificado = True
End Property

Public Property Get RX() As String
    RX = mProps.RX
End Property

Public Property Let TAC(ByVal value As String)
    mProps.TAC = value
    mbooModificado = True
End Property

Public Property Get TAC() As String
    TAC = mProps.TAC
End Property

Public Property Let Otros(ByVal value As String)
    mProps.Otros = value
    mbooModificado = True
End Property

Public Property Get Otros() As String
    Otros = mProps.Otros
End Property

Public Property Let Nota(ByVal value As String)
    mProps.Nota = value
    mbooModificado = True
End Property

Public Property Get Nota() As String
    Nota = mProps.Nota
End Property

Public Property Let ImpDiagnostica(ByVal value As String)
    mProps.ImpDiagnostica = value
    mbooModificado = True
End Property

Public Property Get ImpDiagnostica() As String
    ImpDiagnostica = mProps.ImpDiagnostica
End Property

Public Property Let AptoSinRes(ByVal value As String)
    mProps.AptoSinRes = value
    mbooModificado = True
End Property

Public Property Get AptoSinRes() As String
    AptoSinRes = mProps.AptoSinRes
End Property

Public Property Let AptoConResPer(ByVal value As String)
    mProps.AptoConResPer = value
    mbooModificado = True
End Property

Public Property Get AptoConResPer() As String
    AptoConResPer = mProps.AptoConResPer
End Property

Public Property Let AptoConResLabAdap(ByVal value As String)
    mProps.AptoConResLabAdap = value
    mbooModificado = True
End Property

Public Property Get AptoConResLabAdap() As String
    AptoConResLabAdap = mProps.AptoConResLabAdap
End Property

Public Property Let AptoConResLabRes(ByVal value As String)
    mProps.AptoConResLabRes = value
    mbooModificado = True
End Property

Public Property Get AptoConResLabRes() As String
    AptoConResLabRes = mProps.AptoConResLabRes
End Property

Public Property Let AptoConResObs(ByVal value As String)
    mProps.AptoConResObs = value
    mbooModificado = True
End Property

Public Property Get AptoConResObs() As String
    AptoConResObs = mProps.AptoConResObs
End Property

Public Property Let AptoConResNoApto(ByVal value As String)
    mProps.AptoConResNoApto = value
    mbooModificado = True
End Property

Public Property Get AptoConResNoApto() As String
    AptoConResNoApto = mProps.AptoConResNoApto
End Property

Public Property Let FechaConsVisual(ByVal value As Date)
    mProps.FechaConsVisual = value
    mbooModificado = True
End Property

Public Property Get FechaConsVisual() As Date
    FechaConsVisual = mProps.FechaConsVisual
End Property

Public Property Let FechaConsAudi(ByVal value As Date)
    mProps.FechaConsAudi = value
    mbooModificado = True
End Property

Public Property Get FechaConsAudi() As Date
    FechaConsAudi = mProps.FechaConsAudi
End Property

Public Property Let FechaPresBronco(ByVal value As Date)
    mProps.FechaPresBronco = value
    mbooModificado = True
End Property

Public Property Get FechaPresBronco() As Date
    FechaPresBronco = mProps.FechaPresBronco
End Property

Public Property Let FechaPresToxi(ByVal value As Date)
    mProps.FechaPresToxi = value
    mbooModificado = True
End Property

Public Property Get FechaPresToxi() As Date
    FechaPresToxi = mProps.FechaPresToxi
End Property

Public Property Let FechaPresInf(ByVal value As Date)
    mProps.FechaPresInf = value
    mbooModificado = True
End Property

Public Property Get FechaPresInf() As Date
    FechaPresInf = mProps.FechaPresInf
End Property

Public Property Let FechaExamen(ByVal value As Date)
    mProps.FechaExamen = value
    mbooModificado = True
End Property

Public Property Get FechaExamen() As Date
    FechaExamen = mProps.FechaExamen
End Property

' /* Apto con restricciones - INICIO
Public Property Let CheckAptoConResPer(ByVal value As Boolean)
    mProps.CheckAptoConResPer = value
    mbooModificado = True
End Property

Public Property Get CheckAptoConResPer() As Boolean
    CheckAptoConResPer = mProps.CheckAptoConResPer
End Property

Public Property Let CheckConResLabAdap(ByVal value As Boolean)
    mProps.CheckConResLabAdap = value
    mbooModificado = True
End Property

Public Property Get CheckConResLabAdap() As Boolean
    CheckConResLabAdap = mProps.CheckConResLabAdap
End Property

Public Property Let CheckAptoConResLabRes(ByVal value As Boolean)
    mProps.CheckAptoConResLabRes = value
    mbooModificado = True
End Property

Public Property Get CheckAptoConResLabRes() As Boolean
    CheckAptoConResLabRes = mProps.CheckAptoConResLabRes
End Property

Public Property Let CheckAptoConResObs(ByVal value As Boolean)
    mProps.CheckAptoConResObs = value
    mbooModificado = True
End Property

Public Property Get CheckAptoConResObs() As Boolean
    CheckAptoConResObs = mProps.CheckAptoConResObs
End Property

Public Property Let CheckAptoConResNoApto(ByVal value As Boolean)
    mProps.CheckAptoConResNoApto = value
    mbooModificado = True
End Property

Public Property Get CheckAptoConResNoApto() As Boolean
    CheckAptoConResNoApto = mProps.CheckAptoConResNoApto
End Property

Public Property Let Edad(ByVal value As Integer)
    mProps.Edad = value
    mbooModificado = True
End Property

Public Property Get Edad() As Integer
    Edad = mProps.Edad
End Property
' */ Apto con restricciones - FIN

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

Public Sub Recuperar(ByVal IdEmpleado As Long)
    Dim sql As String, rs As Recordset

    sql = "SELECT * FROM dbo.FichaMedicaExmFisico exm WHERE exm.Idexmfisico =" & IdEmpleado
    Set rs = mProps.mobjEmpleado.Empresa.OpenRecordset(sql)
        
    With rs
        Do While Not .EOF
            mProps.IdExmFisico = !IdExmFisico
            mProps.IdEmpleado = !IdEmpleado
            If Not IsNull(!FechaConsulta) Then mProps.FechaConsulta = !FechaConsulta
            If Not IsNull(!HoraConsulta) Then mProps.HoraConsulta = !HoraConsulta
            If Not IsNull(!MotivoConsulta) Then mProps.MotivoConsulta = !MotivoConsulta
            If Not IsNull(!TipoHistorial) Then mProps.TipoHistorial = !TipoHistorial
            If Not IsNull(!Peso) Then mProps.Peso = !Peso
            If Not IsNull(!TA) Then mProps.TA = !TA
            If Not IsNull(!Talla) Then mProps.Talla = !Talla
            If Not IsNull(!FC) Then mProps.FC = !FC
            If Not IsNull(!temp) Then mProps.temp = !temp
            If Not IsNull(!Biotipo) Then mProps.Biotipo = !Biotipo
            If Not IsNull(!StdNutricional) Then mProps.StdNutricional = !StdNutricional
            If Not IsNull(!AgudezaVisual) Then mProps.AgudezaVisual = !AgudezaVisual
            If Not IsNull(!RejillaAmsler) Then mProps.RejillaAmsler = !RejillaAmsler
            If Not IsNull(!UsaLentes) Then mProps.UsaLentes = !UsaLentes
            If Not IsNull(!Prescripcion) Then mProps.Prescripcion = !Prescripcion
            If Not IsNull(!Receta) Then mProps.Receta = !Receta
            If Not IsNull(!PielFaneras) Then mProps.PielFaneras = !PielFaneras
            If Not IsNull(!Cabeza) Then mProps.Cabeza = !Cabeza
            If Not IsNull(!Ojos) Then mProps.Ojos = !Ojos
            If Not IsNull(!EnfermVista) Then mProps.EnfermVista = !EnfermVista
            If Not IsNull(!Oidos) Then mProps.Oidos = !Oidos
            If Not IsNull(!OtoscopiaOD) Then mProps.OtoscopiaOD = !OtoscopiaOD
            If Not IsNull(!OtoscopiaOI) Then mProps.OtoscopiaOI = !OtoscopiaOI
            If Not IsNull(!AudiometriaOD) Then mProps.AudiometriaOD = !AudiometriaOD
            If Not IsNull(!AudiometriaOI) Then mProps.AudiometriaOI = !AudiometriaOI
            If Not IsNull(!Nariz) Then mProps.Nariz = !Nariz
            If Not IsNull(!Boca) Then mProps.Boca = !Boca
            If Not IsNull(!Dentadura) Then mProps.Dentadura = !Dentadura
            If Not IsNull(!Cuello) Then mProps.Cuello = !Cuello
            If Not IsNull(!Torax) Then mProps.Torax = !Torax
            If Not IsNull(!CsPs) Then mProps.CsPs = !CsPs
            If Not IsNull(!Garganta) Then mProps.Garganta = !Garganta
            If Not IsNull(!Espirometria) Then mProps.Espirometria = !Espirometria
            If Not IsNull(!EKG) Then mProps.EKG = !EKG
            If Not IsNull(!RsCs) Then mProps.RsCs = !RsCs
            If Not IsNull(!PrubEsfuerzo) Then mProps.PrubEsfuerzo = !PrubEsfuerzo
            If Not IsNull(!Abdomen) Then mProps.Abdomen = !Abdomen
            If Not IsNull(!AparatoUro) Then mProps.AparatoUrogenital = !AparatoUro
            '------------
            If Not IsNull(!RegionAno) Then mProps.RegionAno = !RegionAno
            If Not IsNull(!ColumnaVert) Then mProps.ColumnaVert = !ColumnaVert
            If Not IsNull(!EvaPostural) Then mProps.EvaPostural = !EvaPostural
            If Not IsNull(!ExtSuperior) Then mProps.ExtSuperior = !ExtSuperior
            If Not IsNull(!ExtInferiores) Then mProps.ExtInferiores = !ExtInferiores
            If Not IsNull(!Motilidad) Then mProps.Motilidad = !Motilidad
            If Not IsNull(!Sensibilidad) Then mProps.Sensibilidad = !Sensibilidad
            If Not IsNull(!Reflejos) Then mProps.Reflejos = !Reflejos
            If Not IsNull(!Coordinacion) Then mProps.Coordinacion = !Coordinacion
            If Not IsNull(!MovInvol) Then mProps.MovInvol = !MovInvol
            If Not IsNull(!Equilibrio) Then mProps.Equilibrio = !Equilibrio
            If Not IsNull(!MovInvol) Then mProps.MovInvol = !MovInvol
            If Not IsNull(!Conciencia) Then mProps.Conciencia = !Conciencia
            If Not IsNull(!Memoria) Then mProps.Memoria = !Memoria
            If Not IsNull(!GangliosLinf) Then mProps.GangliosLinf = !GangliosLinf
            If Not IsNull(!Sangre) Then mProps.Sangre = !Sangre
            If Not IsNull(!Emo) Then mProps.Emo = !Emo
            If Not IsNull(!Heces) Then mProps.Heces = !Heces
            If Not IsNull(!RX) Then mProps.RX = !RX
            If Not IsNull(!TAC) Then mProps.TAC = !TAC
            If Not IsNull(!Otros) Then mProps.Otros = !Otros
            If Not IsNull(!Nota) Then mProps.Nota = !Nota
            If Not IsNull(!ImpDiagnostica) Then mProps.ImpDiagnostica = !ImpDiagnostica
            If Not IsNull(!AptoSinRes) Then mProps.AptoSinRes = !AptoSinRes
            If Not IsNull(!Nota) Then mProps.Nota = !Nota
            If Not IsNull(!AptoConResPer) Then mProps.AptoConResPer = !AptoConResPer
            If Not IsNull(!AptoConResLabAdap) Then mProps.AptoConResLabAdap = !AptoConResLabAdap
            If Not IsNull(!AptoConResLabRes) Then mProps.AptoConResLabRes = !AptoConResLabRes
            If Not IsNull(!AptoConResObs) Then mProps.AptoConResObs = !AptoConResObs
            If Not IsNull(!AptoConResNoApto) Then mProps.AptoConResNoApto = !AptoConResNoApto
            If Not IsNull(!FechaConsVisual) Then mProps.FechaConsVisual = !FechaConsVisual
            If Not IsNull(!FechaConsAudi) Then mProps.FechaConsAudi = !FechaConsAudi
            If Not IsNull(!FechaPresBronco) Then mProps.FechaPresBronco = !FechaPresBronco
            If Not IsNull(!FechaPresToxi) Then mProps.FechaPresToxi = !FechaPresToxi
            If Not IsNull(!FechaPresInf) Then mProps.FechaPresInf = !FechaPresInf
            If Not IsNull(!FechaExamen) Then mProps.FechaExamen = !FechaExamen
            '/* Consulta Diaria
            If Not IsNull(!ConsultaDiaria) Then mProps.ConsultaDiaria = !ConsultaDiaria
            '*/
            If Not IsNull(!FechaPreocupacional) Then mProps.FechaPreocupacional = !FechaPreocupacional
            If Not IsNull(!FechaOcupacional) Then mProps.FechaOcupacional = !FechaOcupacional
            If Not IsNull(!FechaReintegro) Then mProps.FechaReintegro = !FechaReintegro
            If Not IsNull(!FechaRetiro) Then mProps.FechaRetiro = !FechaRetiro
            ' /* Agudeza y Regilla
            If Not IsNull(!AgudezaVisualOD) Then mProps.AgudezaVisualOD = !AgudezaVisualOD
            If Not IsNull(!AgudezaVisualOI) Then mProps.AgudezaVisualOI = !AgudezaVisualOI
            If Not IsNull(!RejillaAmslerOD) Then mProps.RejillaAmslerOD = !RejillaAmslerOD
            If Not IsNull(!RejillaAmslerOI) Then mProps.RejillaAmslerOI = !RejillaAmslerOI
            
            ' /* Apto con restricciones
            If Not IsNull(!CheckAptoConResPer) Then mProps.CheckAptoConResPer = !CheckAptoConResPer
            If Not IsNull(!CheckConResLabAdap) Then mProps.CheckConResLabAdap = !CheckConResLabAdap
            If Not IsNull(!CheckAptoConResLabRes) Then mProps.CheckAptoConResLabRes = !CheckAptoConResLabRes
            If Not IsNull(!CheckAptoConResObs) Then mProps.CheckAptoConResObs = !CheckAptoConResObs
            If Not IsNull(!CheckAptoConResNoApto) Then mProps.CheckAptoConResNoApto = !CheckAptoConResNoApto
            If Not IsNull(!Edad) Then mProps.Edad = !Edad
            ' */
            If Not IsNull(!RecetaPrescrip) Then mProps.RecetaPrescrip = !RecetaPrescrip
            

            
            rs.MoveNext
        Loop
            .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub RecuperarEnfermedad(ByVal CodEnfermedad As String)
    Dim sql As String, rs As Recordset

    sql = "SELECT * FROM dbo.fichaenfermedad enf WHERE enf.codEnfermedad = '" & CodEnfermedad & "'"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
        
    With rs
        Do While Not .EOF
            mProps.idEnfermedad = !idEnfermedad
            If Not IsNull(!CodEnfermedad) Then mProps.CodEnfermedad = !CodEnfermedad
            If Not IsNull(!Descripcion) Then mProps.DescripcionEnfermedad = !Descripcion
         
            rs.MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Function RecuperarListaTreeView(ByVal DescripcionEnfermedad As String) As Recordset
    Dim sql As String, rs As Recordset

    sql = "SELECT TOP 1000 * FROM dbo.fichaenfermedad WHERE nombrecuenta like '%" & DescripcionEnfermedad & "%' and Bandtotal = 0 ORDER BY CodCuenta ASC"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    Set RecuperarListaTreeView = rs
End Function

Public Function RecuperarListaTreeView01(ByVal IdCuenta As Integer) As Recordset
    Dim sql As String, rs As Recordset

    sql = "SELECT TOP 1 IdCuenta, CodCuenta, NombreCuenta, BandTotal, IdCuentaSuma FROM fichaenfermedad where idCuenta = " & IdCuenta
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    Set RecuperarListaTreeView01 = rs
End Function

Public Sub recuperarEnfermedadUltimo(ByVal CodEnfermedad As String)
    Dim sql As String, rs As Recordset

    sql = "SELECT * FROM dbo.fichaenfermedad enf WHERE enf.codEnfermedad like '" & CodEnfermedad & "%'"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
        
    With rs
        Do While Not .EOF
            mProps.idEnfermedad = !idEnfermedad
            If Not IsNull(!CodEnfermedad) Then mProps.CodEnfermedad = !CodEnfermedad
            rs.MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub ActualizarEnfermedad(ByVal idEnfermedad As Integer, ByVal Descripcion As String)
    Dim sql As String, rs As Recordset
    Dim i As Long
    On Error GoTo Etiq
    
    sql = "UPDATE dbo.FichaEnfermedad SET Descripcion = '" & Descripcion & "' WHERE idEnfermedad = '" & mProps.idEnfermedad & "'"
    Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Public Sub GrabarEnfermedad(ByVal CodEnfermedad As String, ByVal Descripcion As String)
    Dim sql As String, rs As Recordset
    Dim i As Long
    On Error GoTo Etiq
    
    sql = "INSERT INTO dbo.FichaEnfermedad (codEnfermedad, Descripcion, Nivel) VALUES ('" & CodEnfermedad & "', '" & Descripcion & "', 2)"
    Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Public Sub EliminarEnfermedad(ByVal idEnfermedad As String, ByVal Descripcion As String)
    Dim sql As String, rs As Recordset
    Dim i As Long
    On Error GoTo Etiq
    
    sql = "DELETE dbo.FichaEnfermedad WHERE idEnfermedad = " & idEnfermedad
    Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Public Sub Grabar(ByVal IdEmpleado As Long)
    Dim sql As String, rs As Recordset
    Dim i As Long
    On Error GoTo Etiq
        
    sql = "SELECT * FROM FichaMedicaExmFisico WHERE IdExmFisico = " & mProps.IdExmFisico
    Set rs = mProps.mobjEmpleado.Empresa.OpenRecordsetParaEdit(sql)
    
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sql = "SELECT * FROM dbo.FichaMedicaExmFisico WHERE 1=0"
        Set rs = mProps.mobjEmpleado.Empresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    Else
    End If
    
    With rs
            !IdEmpleado = IdEmpleado
'            !IdNroHistorial = mProps.IdNroHistorial
            !FechaConsulta = mProps.FechaConsulta
            !HoraConsulta = mProps.HoraConsulta
            !MotivoConsulta = mProps.MotivoConsulta
            !TipoHistorial = mProps.TipoHistorial
            !Peso = mProps.Peso
            !Talla = mProps.Talla
            !TA = mProps.TA
            !FC = mProps.FC
            !temp = mProps.temp
            !Biotipo = mProps.Biotipo
            !StdNutricional = mProps.StdNutricional
            !AgudezaVisual = mProps.AgudezaVisual
            !RejillaAmsler = mProps.RejillaAmsler
            !UsaLentes = mProps.UsaLentes
            !Prescripcion = mProps.Prescripcion
            !Receta = mProps.Receta
            !PielFaneras = mProps.PielFaneras
            !Cabeza = mProps.Cabeza
            !Ojos = mProps.Ojos
            !EnfermVista = mProps.EnfermVista
            !Oidos = mProps.Oidos
            !OtoscopiaOD = mProps.OtoscopiaOD
            !OtoscopiaOI = mProps.OtoscopiaOI
            !AudiometriaOD = mProps.AudiometriaOD
            !AudiometriaOI = mProps.AudiometriaOI
            !Nariz = mProps.Nariz
            !Boca = mProps.Boca
            !Dentadura = mProps.Dentadura
            !Cuello = mProps.Cuello
            !Torax = mProps.Torax
            !CsPs = mProps.CsPs
            !Garganta = mProps.Garganta
            !Espirometria = mProps.Espirometria
            !RsCs = mProps.RsCs
            !EKG = mProps.EKG
            !PrubEsfuerzo = mProps.PrubEsfuerzo
            !Abdomen = mProps.Abdomen
            !AparatoUro = mProps.AparatoUrogenital
            '----------
            !RegionAno = mProps.RegionAno
            !ColumnaVert = mProps.ColumnaVert
            !EvaPostural = mProps.EvaPostural
            !ExtSuperior = mProps.ExtSuperior
            !ExtInferiores = mProps.ExtInferiores
            !Motilidad = mProps.Motilidad
            !Sensibilidad = mProps.Sensibilidad
            !Reflejos = mProps.Reflejos
            !Coordinacion = mProps.Coordinacion
            !MovInvol = mProps.MovInvol
            !Equilibrio = mProps.Equilibrio
            !MovInvol = mProps.MovInvol
            !Conciencia = mProps.Conciencia
            !Memoria = mProps.Memoria
            !GangliosLinf = mProps.GangliosLinf
            !Sangre = mProps.Sangre
            !Emo = mProps.Emo
            !Heces = mProps.Heces
            !RX = mProps.RX
            !TAC = mProps.TAC
            !Otros = mProps.Otros
            !Nota = mProps.Nota
            !ImpDiagnostica = mProps.ImpDiagnostica
            !AptoSinRes = mProps.AptoSinRes
            !Nota = mProps.Nota
            !AptoConResPer = mProps.AptoConResPer
            !AptoConResLabAdap = mProps.AptoConResLabAdap
            !AptoConResLabRes = mProps.AptoConResLabRes
            !AptoConResObs = mProps.AptoConResObs
            !AptoConResNoApto = mProps.AptoConResNoApto
            !FechaConsVisual = mProps.FechaConsVisual
            !FechaConsAudi = mProps.FechaConsAudi
            !FechaPresBronco = mProps.FechaPresBronco
            !FechaPresToxi = mProps.FechaPresToxi
            !FechaPresInf = mProps.FechaPresInf
            !FechaExamen = mProps.FechaExamen
            '/*
            !ConsultaDiaria = mProps.ConsultaDiaria
            '*/
            !FechaPreocupacional = mProps.FechaPreocupacional
            !FechaOcupacional = mProps.FechaOcupacional
            !FechaReintegro = mProps.FechaReintegro
            !FechaRetiro = mProps.FechaRetiro
            ' /* Agudeza y Rejilla
            !AgudezaVisualOD = mProps.AgudezaVisualOD
            !AgudezaVisualOI = mProps.AgudezaVisualOI
            !RejillaAmslerOD = mProps.RejillaAmslerOD
            !RejillaAmslerOI = mProps.RejillaAmslerOI
            ' /* Apto con restricciones
            !CheckAptoConResPer = mProps.CheckAptoConResPer
            !CheckConResLabAdap = mProps.CheckConResLabAdap
            !CheckAptoConResLabRes = mProps.CheckAptoConResLabRes
            !CheckAptoConResObs = mProps.CheckAptoConResObs
            !CheckAptoConResNoApto = mProps.CheckAptoConResNoApto
            !Edad = mProps.Edad
            
            !RecetaPrescrip = mProps.RecetaPrescrip
            
            
                                    
            .Update
            .Move 0

            mProps.IdExmFisico = !IdExmFisico
        .Close
    End With
    
            sql = "DELETE FROM fichamedenfdetalle WHERE idExmFisico=" & mProps.IdExmFisico
            mProps.mobjEmpleado.Empresa.Execute sql, True
        
        For i = 1 To Me.CountEnfDetalle
             Me.EnfDetalle(i).Grabar
        Next


    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Public Function ListaEnfermedadGrilla(ByVal IdExmFisico As Integer) As Recordset
    Dim sql As String, rs As Recordset
        
    sql = "SELECT codEnfermedad, Descripcion FROM dbo.FichaMedEnfDetalle WHERE IdExmFisico =" & IdExmFisico
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    Set ListaEnfermedadGrilla = rs
    'Set rs = Nothing
End Function

Public Function ListaDetalle() As Variant

End Function


Public Function RecuperaFichaDetalle(key As Variant) As fichaDetalle
    Dim obj As fichaDetalle
    Set obj = New fichaDetalle
    Set obj.objFicha = Me
    On Error Resume Next
    obj.Recuperar key
    If Err.Number = 0 Then Set RecuperaFichaDetalle = obj
    Err.Clear
    On Error GoTo 0
    Set obj = Nothing
End Function

Public Function RecuperaFichaDetallexFicha(key As Variant) As fichaDetalle
    Dim obj As fichaDetalle
    Set obj = New fichaDetalle
    Set obj.objFicha = Me
    On Error Resume Next
    obj.recuperarxFicha key
    If Err.Number = 0 Then Set RecuperaFichaDetallexFicha = obj
    Err.Clear
    On Error GoTo 0
    Set obj = Nothing
End Function


Public Function Eliminar() As Boolean
    Dim sql As String
    Dim rs As Recordset
    On Error GoTo CapturaError
    sql = "DELETE FROM dbo.FichaMedicaExmFisico WHERE idEmpleado = " & Me.IdEmpleado
        
        
    mbooEsNuevo = True
    mbooModificado = True
    Exit Function
    
CapturaError:
    MsgBox Err.Description
    Set rs = Nothing
    Eliminar = False
    Exit Function
End Function


Public Sub RecuperarFechaNacimiento(ByVal IdEmpleado As Long)
   Dim sql As String, rs As Recordset
    sql = "SELECT emp.IdProvCli, emp.FechaNacimiento FROM dbo.Empleado emp WHERE emp.IdProvCli = " & IdEmpleado
    Set rs = mProps.mobjEmpleado.Empresa.OpenRecordset(sql)
    With rs
        Do While Not .EOF
            'mProps.IdEmpleado = !IdEmpleado
            If Not IsNull(!FechaNacimiento) Then mProps.FechaNacimiento = !FechaNacimiento
            rs.MoveNext
        Loop
            .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Property Let FechaNacimiento(ByVal value As Date)
    mProps.FechaNacimiento = value
    mbooModificado = True
End Property

Public Property Get FechaNacimiento() As Date
    FechaNacimiento = mProps.FechaNacimiento
End Property

Public Function AddFichaMedEnf() As Long
    Dim nuevo As FichaMedEnfDetalle
    If mcolEnfDetalle Is Nothing Then RecuperaEnfDetalle
    Set nuevo = New FichaMedEnfDetalle
    Set nuevo.objFichaMedExamen = Me
    mcolEnfDetalle.Add nuevo
    AddFichaMedEnf = mcolEnfDetalle.Count    'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function
Private Sub RecuperaEnfDetalle()
    Dim sql As String, rs As Recordset, i As Long, mod_antes As Boolean
    mod_antes = mbooModificado      'Guarda el estado de Modificado
    If mcolEnfDetalle Is Nothing Then Set mcolEnfDetalle = New Collection
    'BorrarAFKardex
    
    sql = "SELECT * FROM FichaMedEnfDetalle WHERE IdExmFisico=" & mProps.IdExmFisico
    '--& " ORDER BY Orden"
    Set rs = mProps.mobjEmpleado.Empresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            If !IdExmFisDetalle > 0 Then                 'Para asegurar el Bug de ADO
                i = AddFichaMedEnf
                EnfDetalle(i).Recuperar rs
            End If
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    mbooModificado = mod_antes      'Para que no cambie el estado de Modificado
End Sub

Public Property Get CountEnfDetalle() As Long
    If mcolEnfDetalle Is Nothing Then RecuperaEnfDetalle
    CountEnfDetalle = mcolEnfDetalle.Count
End Property

Public Property Get EnfDetalle(ix As Long) As FichaMedEnfDetalle
    If mcolEnfDetalle Is Nothing Then RecuperaEnfDetalle
    Set EnfDetalle = mcolEnfDetalle.item(ix)
End Property

Public Sub RemoveEnfDetalle(ByVal ix As Long, Optional obj As FichaMedEnfDetalle)
If mcolEnfDetalle Is Nothing Then RecuperaEnfDetalle
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolEnfDetalle.Count
            If mcolEnfDetalle.item(ix) Is obj Then
                mcolEnfDetalle.Remove ix
                Exit For
            End If
        Next ix
        'Debug.Assert True   'No se puede dar
    Else
        mcolEnfDetalle.Remove ix
    End If
    mbooModificado = True
End Sub

Public Sub BorradoFisico(cod As Long, id As Long)
    Dim sql As String
    Dim rs As Recordset
    On Error GoTo CapturaError
    
    sql = "delete from fichaMedEnfDetalle where IdExmfisico= " & id
    Me.mobjEmpleado.Empresa.EjecutarSQL sql, 1
    
    sql = "delete from fichamedicaexmfisico where IdExmfisico= " & id
    Me.mobjEmpleado.Empresa.EjecutarSQL sql, 1
    
    Set rs = Nothing
    Exit Sub
CapturaError:
'    MsgBox Err.Description
    Set rs = Nothing
    Exit Sub

        
End Sub


Public Property Let RecetaPrescrip(ByVal value As String)
    mProps.RecetaPrescrip = value
    mbooModificado = True
End Property

Public Property Get RecetaPrescrip() As String
    RecetaPrescrip = mProps.RecetaPrescrip
End Property

