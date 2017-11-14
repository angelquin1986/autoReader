VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FichaMedicaEmp"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdEmpleado As Long
    Instruccion As String
    Profesion   As String
    Ocupacion  As String
    IdPais As Long
    CodPais As String
    IdCanton    As Long
    CodCanton As String
    IDParroquia As Long
    IdProvincia As Long
    CodParroquia As String
    CodProvincia As String
    RecidOcasional As String
    RecidActual  As String
    Religion As String
    AntePersonales As String
    Vacunacion As String
    Cirugias As String
    Alergias As String
    TomaMedicacion  As Boolean
    DetalleMedicacion  As String
    Tabaco As Boolean
    TabacoDia As Integer
    Alcohol  As Boolean
    AlcoholSemana  As Integer
    Drogas   As String
    Medicamentos As String
    Menarquia As String
    Menopausia  As String
    FUM As String
    Ciclos          As String
    AntConceptivoG As String
    AntConceptivoP As String
    AntConceptivoHV    As String
    AntConceptivoA  As String
    Familiares As String
    UsoAnticonceptivos As String
    
    TabacoNota As String
    AlcoholNota As String
    Droga As Boolean
    NombreEmergencia As String
    TelefonoEmergencia As String
    RelacionEmergencia As String
    
    Zona As String
    TipoDiscapacidad As Integer
    TieneCarnetDiscap As Integer
    CarnetDiscapActualiza As Integer
    FechaActualizaCarnet As Date
    NotaTipoDiscapacidad As String
    NombreEmergencia2 As String
    TelefonoEmergencia2 As String
    RelacionEmergencia2 As String
    NombreEmergencia3 As String
    TelefonoEmergencia3 As String
    RelacionEmergencia3 As String
    Anteriores2 As String
    Anteriores3 As String
    AnosAnt As Integer
    AnosAnt2 As Integer
    AnosAnt3 As Integer
    RiesgoAnt As String
    RiesgoAnt2 As String
    RiesgoAnt3 As String
    ProtexionAnt As Boolean
    ProtexionAnt2 As Boolean
    ProtexionAnt3 As Boolean
    
    
    mobjEmpleado As PCProvCli
    
    
    Telefono As String
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

Public Property Let Instruccion(ByVal value As String)
    mProps.Instruccion = value
    mbooModificado = True
End Property

Public Property Get Instruccion() As String
    Instruccion = mProps.Instruccion
End Property

Public Property Let Telefono(ByVal value As String)
    mProps.Telefono = value
    mbooModificado = True
End Property

Public Property Get Telefono() As String
    Telefono = mProps.Telefono
End Property

Public Property Let UsoAnticonceptivos(ByVal value As String)
    mProps.UsoAnticonceptivos = value
    mbooModificado = True
End Property

Public Property Get UsoAnticonceptivos() As String
    UsoAnticonceptivos = mProps.UsoAnticonceptivos
End Property

Public Property Let Profesion(ByVal value As String)
    mProps.Profesion = value
    mbooModificado = True
End Property

Public Property Get Profesion() As String
    Profesion = mProps.Profesion
End Property

Public Property Let Ocupacion(ByVal value As String)
    mProps.Ocupacion = value
    mbooModificado = True
End Property

Public Property Get Ocupacion() As String
    Ocupacion = mProps.Ocupacion
End Property

Public Property Let IdPais(ByVal value As Long)
    mProps.IdPais = value
    mbooModificado = True
End Property

Public Property Get IdPais() As Long
    IdPais = mProps.IdPais
End Property

Public Property Let CodPais(ByVal value As String)
Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodPais Then
        If Len(value) > 0 Then
            sql = "SELECT Id FROM Anexo_Pais WHERE Codpais='" & value & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Anexo_Pais_CodPais", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdPais = rs!id
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdPais = 0
        End If
    End If

    mProps.CodPais = value
    mbooModificado = True
End Property

Public Property Get CodPais() As String
    CodPais = mProps.CodPais
End Property

Public Property Let IdCanton(ByVal value As Long)
    mProps.IdCanton = value
    mbooModificado = True
End Property

Public Property Get IdCanton() As Long
    IdCanton = mProps.IdCanton
End Property

Public Property Let CodCanton(ByVal value As String)
Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodCanton Then
        If Len(value) > 0 Then
            sql = "SELECT Idcanton FROM PCCanton WHERE Codcanton='" & value & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Canton_Codcanton", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCanton = rs!IdCanton
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdCanton = 0
        End If
    End If
    mProps.CodCanton = value
    mbooModificado = True
End Property

Public Property Get CodCanton() As String
    CodCanton = mProps.CodCanton
End Property

Public Property Let IDParroquia(ByVal value As Long)
    mProps.IDParroquia = value
    mbooModificado = True
End Property

Public Property Get IDParroquia() As Long
    IDParroquia = mProps.IDParroquia
End Property

Public Property Let IdProvincia(ByVal value As Long)
    mProps.IdProvincia = value
    mbooModificado = True
End Property

Public Property Get IdProvincia() As Long
    IdProvincia = mProps.IdProvincia
End Property

Public Property Let CodProvincia(ByVal value As String)
Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodParroquia Then
        If Len(value) > 0 Then
            sql = "SELECT IdProvincia FROM dbo.PCProvincia WHERE CodProvincia ='" & value & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Provincia_CodProvincia", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdProvincia = rs!IdProvincia
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdProvincia = 0
        End If
    End If

    mProps.CodParroquia = value
    mbooModificado = True
End Property

Public Property Get CodProvincia() As String
    CodProvincia = mProps.CodProvincia
End Property

Public Property Let CodParroquia(ByVal value As String)
Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodParroquia Then
        If Len(value) > 0 Then
            sql = "SELECT IdParroquia FROM PCParroquia WHERE CodParroquia='" & value & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Parroquia_CodParroquia", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IDParroquia = rs!IDParroquia
            End If
            rs.Close
            Set rs = Nothing
        Else
            IDParroquia = 0
        End If
    End If

    mProps.CodParroquia = value
    mbooModificado = True
End Property

Public Property Get CodParroquia() As String
    CodParroquia = mProps.CodParroquia
End Property
Public Property Let RecidOcasional(ByVal value As String)
    mProps.RecidOcasional = value
    mbooModificado = True
End Property

Public Property Get RecidOcasional() As String
    RecidOcasional = mProps.RecidOcasional
End Property

Public Property Let RecidActual(ByVal value As String)
    mProps.RecidActual = value
    mbooModificado = True
End Property

Public Property Get RecidActual() As String
    RecidActual = mProps.RecidActual
End Property

Public Property Let Religion(ByVal value As String)
    mProps.Religion = value
    mbooModificado = True
End Property

Public Property Get Religion() As String
    Religion = mProps.Religion
End Property

Public Property Let AntePersonales(ByVal value As String)
    mProps.AntePersonales = value
    mbooModificado = True
End Property

Public Property Get AntePersonales() As String
    AntePersonales = mProps.AntePersonales
End Property

Public Property Let Vacunacion(ByVal value As String)
    mProps.Vacunacion = value
    mbooModificado = True
End Property

Public Property Get Vacunacion() As String
    Vacunacion = mProps.Vacunacion
End Property

Public Property Let Cirugias(ByVal value As String)
    mProps.Cirugias = value
    mbooModificado = True
End Property

Public Property Get Cirugias() As String
    Cirugias = mProps.Cirugias
End Property

Public Property Let Alergias(ByVal value As String)
    mProps.Alergias = value
    mbooModificado = True
End Property

Public Property Get Alergias() As String
    Alergias = mProps.Alergias
End Property

Public Property Let TomaMedicacion(ByVal value As Boolean)
    mProps.TomaMedicacion = value
    mbooModificado = True
End Property

Public Property Get TomaMedicacion() As Boolean
    TomaMedicacion = mProps.TomaMedicacion
End Property

Public Property Let DetalleMedicacion(ByVal value As String)
    mProps.DetalleMedicacion = value
    mbooModificado = True
End Property

Public Property Get DetalleMedicacion() As String
    DetalleMedicacion = mProps.DetalleMedicacion
End Property

Public Property Let Tabaco(ByVal value As Boolean)
    mProps.Tabaco = value
    mbooModificado = True
End Property

Public Property Get Tabaco() As Boolean
    Tabaco = mProps.Tabaco
End Property

Public Property Let TabacoDia(ByVal value As Integer)
    mProps.TabacoDia = value
    mbooModificado = True
End Property

Public Property Get TabacoDia() As Integer
    TabacoDia = mProps.TabacoDia
End Property

Public Property Let Alcohol(ByVal value As Boolean)
    mProps.Alcohol = value
    mbooModificado = True
End Property

Public Property Get Alcohol() As Boolean
    Alcohol = mProps.Alcohol
End Property

Public Property Let AlcoholSemana(ByVal value As Integer)
    mProps.AlcoholSemana = value
    mbooModificado = True
End Property

Public Property Get AlcoholSemana() As Integer
    AlcoholSemana = mProps.AlcoholSemana
End Property

Public Property Let Drogas(ByVal value As String)
    mProps.Drogas = value
    mbooModificado = True
End Property

Public Property Get Drogas() As String
    Drogas = mProps.Drogas
End Property

Public Property Let Medicamentos(ByVal value As String)
    mProps.Medicamentos = value
    mbooModificado = True
End Property

Public Property Get Medicamentos() As String
    Medicamentos = mProps.Medicamentos
End Property

Public Property Let Menarquia(ByVal value As String)
    mProps.Menarquia = value
    mbooModificado = True
End Property

Public Property Get Menarquia() As String
    Menarquia = mProps.Menarquia
End Property

Public Property Let Menopausia(ByVal value As String)
    mProps.Menopausia = value
    mbooModificado = True
End Property

Public Property Get Menopausia() As String
    Menopausia = mProps.Menopausia
End Property

Public Property Let FUM(ByVal value As String)
    mProps.FUM = value
    mbooModificado = True
End Property

Public Property Get FUM() As String
    FUM = mProps.FUM
End Property

Public Property Let Ciclos(ByVal value As String)
    mProps.Ciclos = value
    mbooModificado = True
End Property

Public Property Get Ciclos() As String
    Ciclos = mProps.Ciclos
End Property

Public Property Let AntConceptivoG(ByVal value As String)
    mProps.AntConceptivoG = value
    mbooModificado = True
End Property

Public Property Get AntConceptivoG() As String
    AntConceptivoG = mProps.AntConceptivoG
End Property

Public Property Let AntConceptivoP(ByVal value As String)
    mProps.AntConceptivoP = value
    mbooModificado = True
End Property

Public Property Get AntConceptivoP() As String
    AntConceptivoP = mProps.AntConceptivoP
End Property

Public Property Let AntConceptivoHV(ByVal value As String)
    mProps.AntConceptivoHV = value
    mbooModificado = True
End Property

Public Property Get AntConceptivoHV() As String
    AntConceptivoHV = mProps.AntConceptivoHV
End Property

Public Property Let Familiares(ByVal value As String)
    mProps.Familiares = value
    mbooModificado = True
End Property

Public Property Get Familiares() As String
    Familiares = mProps.Familiares
End Property

Public Property Let AntConceptivoA(ByVal value As String)
    mProps.AntConceptivoA = value
    mbooModificado = True
End Property

Public Property Get AntConceptivoA() As String
    AntConceptivoA = mProps.AntConceptivoA
End Property

Public Property Get IdEmpleado() As Long
    IdEmpleado = mProps.IdEmpleado
End Property

Public Property Let IdEmpleado(ByVal value As Long)
    mProps.IdEmpleado = value
    mbooModificado = True
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
    
'    sql = "SELECT fichamedicaemp.*,PCProvincia.CodProvincia,pccanton.codcanton,pcparroquia.codparroquia,anexo_pais.codpais "
'    sql = sql & " from fichamedicaemp Left Join PCcanton on PCcanton.idcanton = fichamedicaemp.idcanton  "
'    sql = sql & " Left Join PCparroquia on PCparroquia.idparroquia= fichamedicaemp.idParroquia  "
'    sql = sql & " Left Join PCProvincia on PCProvincia.idprovincia= fichamedicaemp.idprovincia  "
'    sql = sql & " Left Join Anexo_pais on Anexo_pais.id= fichamedicaemp.idpais "
'    sql = sql & "WHERE idEmpleado = " & Id
        
    sql = "SELECT fichamedicaemp.*,pccanton.codcanton,pcparroquia.codparroquia,anexo_pais.codpais "
    sql = sql & " from fichamedicaemp Left Join PCcanton on PCcanton.idcanton = fichamedicaemp.idcanton  "
    sql = sql & " Left Join PCparroquia on PCparroquia.idparroquia= fichamedicaemp.idParroquia  "
    sql = sql & " Left Join Anexo_pais on Anexo_pais.id= fichamedicaemp.idpais "
    sql = sql & "WHERE idEmpleado = " & id
    Set rs = mProps.mobjEmpleado.Empresa.OpenRecordset(sql)
    With rs
        Do While Not .EOF
            mProps.IdEmpleado = !IdEmpleado
            If Not IsNull(!Instruccion) Then mProps.Instruccion = !Instruccion
            If Not IsNull(!Profesion) Then mProps.Profesion = !Profesion
            If Not IsNull(!Ocupacion) Then mProps.Ocupacion = !Ocupacion
            mProps.IdPais = !IdPais
            If Not IsNull(!CodPais) Then mProps.CodPais = !CodPais
            mProps.IdCanton = !IdCanton
            If Not IsNull(!CodCanton) Then mProps.CodCanton = !CodCanton
            mProps.IDParroquia = !IDParroquia
            'mProps.IdProvincia = !IdProvincia
            If Not IsNull(!CodParroquia) Then mProps.CodParroquia = !CodParroquia
            If Not IsNull(!RecidOcasional) Then mProps.RecidOcasional = !RecidOcasional
            If Not IsNull(!RecidActual) Then mProps.RecidActual = !RecidActual
            If Not IsNull(!Religion) Then mProps.Religion = !Religion
            If Not IsNull(!Personales) Then mProps.AntePersonales = !Personales
            If Not IsNull(!Vacunacion) Then mProps.Vacunacion = !Vacunacion
            If Not IsNull(!Cirugias) Then mProps.Cirugias = !Cirugias
            If Not IsNull(!Alergias) Then mProps.Alergias = !Alergias
            
            If Not IsNull(!TomaMedicacion) Then mProps.TomaMedicacion = !TomaMedicacion
            If Not IsNull(!DetalleMedicacion) Then mProps.DetalleMedicacion = !DetalleMedicacion
            If Not IsNull(!Tabaco) Then mProps.Tabaco = !Tabaco
            If Not IsNull(!TabacoDia) Then mProps.TabacoDia = !TabacoDia
            If Not IsNull(!Alcohol) Then mProps.Alcohol = !Alcohol
            If Not IsNull(!AlcoholSemana) Then mProps.AlcoholSemana = !AlcoholSemana
            If Not IsNull(!Drogas) Then mProps.Drogas = !Drogas
            If Not IsNull(!Medicamentos) Then mProps.Medicamentos = !Medicamentos
            If Not IsNull(!Menarquia) Then mProps.Menarquia = !Menarquia
            If Not IsNull(!Menopausia) Then mProps.Menopausia = !Menopausia
            If Not IsNull(!FUM) Then mProps.FUM = !FUM
            If Not IsNull(!Ciclos) Then mProps.Ciclos = !Ciclos
            If Not IsNull(!AntConceptivoG) Then mProps.AntConceptivoG = !AntConceptivoG
            If Not IsNull(!AntConceptivoP) Then mProps.AntConceptivoP = !AntConceptivoP
            If Not IsNull(!AntConceptivoHV) Then mProps.AntConceptivoHV = !AntConceptivoHV
            If Not IsNull(!AntConceptivoA) Then mProps.AntConceptivoA = !AntConceptivoA
            If Not IsNull(!Familiares) Then mProps.Familiares = !Familiares
            If Not IsNull(!UsoAnticonceptivos) Then mProps.UsoAnticonceptivos = !UsoAnticonceptivos
                        
            If Not IsNull(!TabacoNota) Then mProps.TabacoNota = !TabacoNota
            If Not IsNull(!AlcoholNota) Then mProps.AlcoholNota = !AlcoholNota
            If Not IsNull(!Droga) Then mProps.Droga = !Droga
            If Not IsNull(!NombreEmergencia) Then mProps.NombreEmergencia = !NombreEmergencia
            If Not IsNull(!TelefonoEmergencia) Then mProps.TelefonoEmergencia = !TelefonoEmergencia
            If Not IsNull(!RelacionEmergencia) Then mProps.RelacionEmergencia = !RelacionEmergencia
            If Not IsNull(!Zona) Then mProps.Zona = !Zona
            If Not IsNull(!TipoDiscapacidad) Then mProps.TipoDiscapacidad = !TipoDiscapacidad
            If Not IsNull(!TieneCarnetDiscap) Then mProps.TieneCarnetDiscap = !TieneCarnetDiscap
            If Not IsNull(!CarnetDiscapActualiza) Then mProps.CarnetDiscapActualiza = !CarnetDiscapActualiza
            If Not IsNull(!FechaActualizaCarnet) Then mProps.FechaActualizaCarnet = !FechaActualizaCarnet
            If Not IsNull(!NotaTipoDiscapacidad) Then mProps.NotaTipoDiscapacidad = !NotaTipoDiscapacidad
            If Not IsNull(!NombreEmergencia2) Then mProps.NombreEmergencia2 = !NombreEmergencia2
            If Not IsNull(!NombreEmergencia3) Then mProps.NombreEmergencia3 = !NombreEmergencia3
            If Not IsNull(!TelefonoEmergencia2) Then mProps.TelefonoEmergencia2 = !TelefonoEmergencia2
            If Not IsNull(!TelefonoEmergencia3) Then mProps.TelefonoEmergencia3 = !TelefonoEmergencia3
            If Not IsNull(!RelacionEmergencia2) Then mProps.RelacionEmergencia2 = !RelacionEmergencia2
            If Not IsNull(!RelacionEmergencia3) Then mProps.RelacionEmergencia3 = !RelacionEmergencia3
            If Not IsNull(!Anteriores2) Then mProps.Anteriores2 = !Anteriores2
            If Not IsNull(!Anteriores3) Then mProps.Anteriores3 = !Anteriores3
            If Not IsNull(!AnosAnt) Then mProps.AnosAnt = !AnosAnt
            If Not IsNull(!AnosAnt2) Then mProps.AnosAnt2 = !AnosAnt2
            If Not IsNull(!AnosAnt3) Then mProps.AnosAnt3 = !AnosAnt3
            If Not IsNull(!RiesgoAnt) Then mProps.RiesgoAnt = !RiesgoAnt
            If Not IsNull(!RiesgoAnt2) Then mProps.RiesgoAnt2 = !RiesgoAnt2
            If Not IsNull(!RiesgoAnt3) Then mProps.RiesgoAnt3 = !RiesgoAnt3
            If Not IsNull(!ProtexionAnt) Then mProps.ProtexionAnt = !ProtexionAnt
            If Not IsNull(!ProtexionAnt2) Then mProps.ProtexionAnt2 = !ProtexionAnt2
            If Not IsNull(!ProtexionAnt3) Then mProps.ProtexionAnt3 = !ProtexionAnt3
            
            
                        
                        
            rs.MoveNext
        Loop
            .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    RecuperarLugarNacimiento (id)
End Sub

Public Sub RecuperarLugarNacimiento(ByVal id As Long)
    Dim sql As String, rs As Recordset
    sql = "SELECT emd.idProvincia, emd.idCanton, emd.idParroquia FROM empleado emd WHERE emd.IdProvCli = " & id
    Set rs = mProps.mobjEmpleado.Empresa.OpenRecordset(sql)
    
    With rs
        Do While Not .EOF
            mProps.IDParroquia = !IDParroquia
            mProps.IdProvincia = !IdProvincia
            mProps.IdCanton = !IdCanton
             
            rs.MoveNext
        Loop
            .Close
    End With
    
    '// Optener Codigo de canton
        If mProps.IdCanton > 0 Then
            sql = "SELECT CodCanton FROM PCCanton WHERE IdCanton='" & mProps.IdCanton & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Canton_IdCanton", MSGERR_NOHAYCODIGO
                Exit Sub
            Else
                mProps.CodCanton = rs!CodCanton
            End If
            rs.Close
            Set rs = Nothing
        Else
            mProps.CodCanton = 0
        End If
    
    '// Optener Codigo de Parroquia
        If mProps.IDParroquia > 0 Then
            sql = "SELECT CodParroquia FROM PCParroquia WHERE IdParroquia='" & mProps.IDParroquia & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Canton_IdParroquia", MSGERR_NOHAYCODIGO
                Exit Sub
            Else
                mProps.CodParroquia = rs!CodParroquia
            End If
            rs.Close
            Set rs = Nothing
        Else
            mProps.CodParroquia = 0
        End If
        
        '// Optener Codigo de provincia
        If mProps.IdProvincia > 0 Then
            sql = "SELECT CodProvincia FROM PCprovincia WHERE IdProvincia='" & mProps.IdProvincia & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Canton_IdProvincia", MSGERR_NOHAYCODIGO
                Exit Sub
            Else
                mProps.CodProvincia = rs!CodProvincia
            End If
            rs.Close
            Set rs = Nothing
        Else
            mProps.CodCanton = 0
        End If
    mbooModificado = True
        
End Sub

Public Sub Grabar(ByVal IdEmpleado As Long)
    Dim sql As String, rs As Recordset
    On Error GoTo Etiq
    
 
            sql = "SELECT * FROM fichamedicaemp WHERE idempleado = " & IdEmpleado
            Set rs = mProps.mobjEmpleado.Empresa.OpenRecordsetParaEdit(sql)
        
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            sql = "SELECT * FROM fichamedicaemp WHERE 1=0"
            Set rs = mProps.mobjEmpleado.Empresa.OpenRecordsetParaEdit(sql)
            rs.AddNew
        Else
        End If
        
    With rs
            !IdEmpleado = IdEmpleado
            !Instruccion = mProps.Instruccion
            !Profesion = mProps.Profesion
            !Ocupacion = mProps.Ocupacion
             !IdPais = mProps.IdPais
            !IdCanton = mProps.IdCanton
            !IDParroquia = mProps.IDParroquia
            '!IdProvincia = mProps.IdProvincia
            !RecidOcasional = mProps.RecidOcasional
            !RecidActual = mProps.RecidActual
            !Religion = mProps.Religion
            !Personales = mProps.AntePersonales
            !Vacunacion = mProps.Vacunacion
            !Cirugias = mProps.Cirugias
            !Alergias = mProps.Alergias
            !TomaMedicacion = mProps.TomaMedicacion
            !DetalleMedicacion = mProps.DetalleMedicacion
            !Tabaco = mProps.Tabaco
            !TabacoDia = mProps.TabacoDia
            !Alcohol = mProps.Alcohol
            !AlcoholSemana = mProps.AlcoholSemana
            !Drogas = mProps.Drogas
            !Medicamentos = mProps.Medicamentos
            !Menarquia = mProps.Menarquia
             !Menopausia = mProps.Menopausia
            !FUM = mProps.FUM
            !Ciclos = mProps.Ciclos
            !AntConceptivoG = mProps.AntConceptivoG
            !AntConceptivoP = mProps.AntConceptivoP
            !AntConceptivoHV = mProps.AntConceptivoHV
            !AntConceptivoA = mProps.AntConceptivoA
            !Familiares = mProps.Familiares
            !UsoAnticonceptivos = mProps.UsoAnticonceptivos
            !TabacoNota = mProps.TabacoNota
            !AlcoholNota = mProps.AlcoholNota
            !Droga = mProps.Droga
            !NombreEmergencia = mProps.NombreEmergencia
            !TelefonoEmergencia = mProps.TelefonoEmergencia
            !RelacionEmergencia = mProps.RelacionEmergencia
            !Zona = mProps.Zona
            
            !TipoDiscapacidad = mProps.TipoDiscapacidad
            !TieneCarnetDiscap = mProps.TieneCarnetDiscap
            !CarnetDiscapActualiza = mProps.CarnetDiscapActualiza
            !FechaActualizaCarnet = mProps.FechaActualizaCarnet
            !NotaTipoDiscapacidad = mProps.NotaTipoDiscapacidad
            !NombreEmergencia2 = mProps.NombreEmergencia2
            !NombreEmergencia3 = mProps.NombreEmergencia3
            !TelefonoEmergencia2 = mProps.TelefonoEmergencia2
            !TelefonoEmergencia3 = mProps.TelefonoEmergencia3
            !RelacionEmergencia2 = mProps.RelacionEmergencia2
            !RelacionEmergencia3 = mProps.RelacionEmergencia3
            !Anteriores2 = mProps.Anteriores2
            !Anteriores3 = mProps.Anteriores3
            !AnosAnt = mProps.AnosAnt
            !AnosAnt2 = mProps.AnosAnt2
            !AnosAnt3 = mProps.AnosAnt3
            !RiesgoAnt = mProps.RiesgoAnt
            !RiesgoAnt2 = mProps.RiesgoAnt2
            !RiesgoAnt3 = mProps.RiesgoAnt3
            !ProtexionAnt = mProps.ProtexionAnt
            !ProtexionAnt2 = mProps.ProtexionAnt2
            !ProtexionAnt3 = mProps.ProtexionAnt3
            
            
            
        .Update
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    GrabarLugarNacimiento (IdEmpleado)
    GrabarTelefono (IdEmpleado)
    
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Public Sub GrabarTelefono(ByVal IdEmpleado As Long)
    Dim sql As String, rs As Recordset
    On Error GoTo Etiq
     
    sql = "SELECT * FROM empleado WHERE IdProvCli = " & IdEmpleado
    Set rs = mProps.mobjEmpleado.Empresa.OpenRecordsetParaEdit(sql)
    
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sql = "SELECT * FROM empleado WHERE 1=0"
        Set rs = mProps.mobjEmpleado.Empresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    End If
    
    With rs
        !Telefono1 = mProps.Telefono
        .Update
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Public Sub GrabarLugarNacimiento(ByVal IdEmpleado As Long)
    Dim sql As String, rs As Recordset
    On Error GoTo Etiq
     
    sql = "SELECT * FROM empleado WHERE IdProvCli = " & IdEmpleado
    Set rs = mProps.mobjEmpleado.Empresa.OpenRecordsetParaEdit(sql)
    
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sql = "SELECT * FROM empleado WHERE 1=0"
        Set rs = mProps.mobjEmpleado.Empresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    End If
    
    With rs
        '!IdProvCli = idEmpleado
        !IdProvincia = mProps.IdProvincia
        !IdCanton = mProps.IdCanton
        !IDParroquia = mProps.IDParroquia
        
        .Update
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    
        sql = "DELETE FROM fichamedicaemp WHERE idEmpleado = " & Me.IdEmpleado
        
    mbooEsNuevo = True
    mbooModificado = True
    
End Sub
'

Public Function ListaDetalleEmpleado() As Variant
    Dim sql As String, rs As Recordset
    sql = "SELECT idexmfisico as 'Id Examen',fechaConsulta as 'Fecha Consulta',horaConsulta as 'Hora Consulta',"
    'motivoconsulta as 'Motivo Consulta',
    sql = sql & " case TipoHistorial  when 0 then 'Preocupacional'  when 1 then 'Ocupacional'  when 2 then 'Reintegro' when 3 then 'Retiro'  else 'Otros' end as 'MotivoConsulta', "
    sql = sql & " prescripcion as 'Prescripción',receta as '-'"
    sql = sql & " FROM FichaMedicaExmFisico  "
    sql = sql & " Where IdEmpleado = " & mProps.IdEmpleado
    sql = sql & " and ConsultaDiaria <> 1 ORDER BY idexmfisico"
    Set rs = Me.mobjEmpleado.Empresa.OpenRecordset(sql)
    Set ListaDetalleEmpleado = rs
    Set rs = Nothing
End Function

Public Function ListaDetalleEmpleado2() As Variant
    Dim sql As String, rs As Recordset
    sql = "SELECT idexmfisico as 'Id Examen',fechaConsulta as 'Fecha Consulta',horaConsulta as 'Hora Consulta',prescripcion as Diagnostico,receta as Prescripción"
    sql = sql & " FROM FichaMedicaExmFisico  "
    sql = sql & " Where IdEmpleado = " & mProps.IdEmpleado
    sql = sql & " and ConsultaDiaria <> 0 ORDER BY idexmfisico"
    Set rs = Me.mobjEmpleado.Empresa.OpenRecordset(sql)
    Set ListaDetalleEmpleado2 = rs
    Set rs = Nothing
End Function

Public Property Let Droga(ByVal value As Boolean)
    mProps.Droga = value
    mbooModificado = True
End Property

Public Property Get Droga() As Boolean
    Droga = mProps.Droga
End Property


Public Property Let TabacoNota(ByVal value As String)
    mProps.TabacoNota = value
    mbooModificado = True
End Property

Public Property Get TabacoNota() As String
    TabacoNota = mProps.TabacoNota
End Property

Public Property Let AlcoholNota(ByVal value As String)
    mProps.AlcoholNota = value
    mbooModificado = True
End Property

Public Property Get AlcoholNota() As String
    AlcoholNota = mProps.AlcoholNota
End Property

Public Property Let NombreEmergencia(ByVal value As String)
    mProps.NombreEmergencia = value
    mbooModificado = True
End Property

Public Property Get NombreEmergencia() As String
    NombreEmergencia = mProps.NombreEmergencia
End Property

Public Property Let TelefonoEmergencia(ByVal value As String)
    mProps.TelefonoEmergencia = value
    mbooModificado = True
End Property

Public Property Get TelefonoEmergencia() As String
    TelefonoEmergencia = mProps.TelefonoEmergencia
End Property

Public Property Let RelacionEmergencia(ByVal value As String)
    mProps.RelacionEmergencia = value
    mbooModificado = True
End Property

Public Property Get RelacionEmergencia() As String
    RelacionEmergencia = mProps.RelacionEmergencia
End Property

Public Property Let Zona(ByVal value As String)
    mProps.Zona = value
    mbooModificado = True
End Property

Public Property Get Zona() As String
    Zona = mProps.Zona
End Property

Public Property Let NotaTipoDiscapacidad(ByVal value As String)
    mProps.NotaTipoDiscapacidad = value
    mbooModificado = True
End Property

Public Property Get NotaTipoDiscapacidad() As String
    NotaTipoDiscapacidad = mProps.NotaTipoDiscapacidad
End Property

Public Property Let NombreEmergencia2(ByVal value As String)
    mProps.NombreEmergencia2 = value
    mbooModificado = True
End Property

Public Property Get NombreEmergencia2() As String
    NombreEmergencia2 = mProps.NombreEmergencia2
End Property

Public Property Let NombreEmergencia3(ByVal value As String)
    mProps.NombreEmergencia3 = value
    mbooModificado = True
End Property

Public Property Get NombreEmergencia3() As String
    NombreEmergencia3 = mProps.NombreEmergencia3
End Property


Public Property Let TelefonoEmergencia2(ByVal value As String)
    mProps.TelefonoEmergencia2 = value
    mbooModificado = True
End Property

Public Property Get TelefonoEmergencia2() As String
    TelefonoEmergencia2 = mProps.TelefonoEmergencia2
End Property

Public Property Let TelefonoEmergencia3(ByVal value As String)
    mProps.TelefonoEmergencia3 = value
    mbooModificado = True
End Property

Public Property Get TelefonoEmergencia3() As String
    TelefonoEmergencia3 = mProps.TelefonoEmergencia3
End Property

Public Property Let RelacionEmergencia2(ByVal value As String)
    mProps.RelacionEmergencia2 = value
    mbooModificado = True
End Property

Public Property Get RelacionEmergencia2() As String
    RelacionEmergencia2 = mProps.RelacionEmergencia2
End Property

Public Property Let RelacionEmergencia3(ByVal value As String)
    mProps.RelacionEmergencia3 = value
    mbooModificado = True
End Property

Public Property Get RelacionEmergencia3() As String
    RelacionEmergencia3 = mProps.RelacionEmergencia3
End Property

Public Property Let Anteriores2(ByVal value As String)
    mProps.Anteriores2 = value
    mbooModificado = True
End Property

Public Property Get Anteriores2() As String
    Anteriores2 = mProps.Anteriores2
End Property

Public Property Let Anteriores3(ByVal value As String)
    mProps.Anteriores3 = value
    mbooModificado = True
End Property

Public Property Get Anteriores3() As String
    Anteriores3 = mProps.Anteriores3
End Property

Public Property Let RiesgoAnt(ByVal value As String)
    mProps.RiesgoAnt = value
    mbooModificado = True
End Property

Public Property Get RiesgoAnt() As String
    RiesgoAnt = mProps.RiesgoAnt
End Property

Public Property Let RiesgoAnt2(ByVal value As String)
    mProps.RiesgoAnt2 = value
    mbooModificado = True
End Property

Public Property Get RiesgoAnt2() As String
    RiesgoAnt2 = mProps.RiesgoAnt2
End Property

Public Property Let RiesgoAnt3(ByVal value As String)
    mProps.RiesgoAnt3 = value
    mbooModificado = True
End Property

Public Property Get RiesgoAnt3() As String
    RiesgoAnt3 = mProps.RiesgoAnt3
End Property

Public Property Let AnosAnt(ByVal value As Integer)
    mProps.AnosAnt = value
    mbooModificado = True
End Property

Public Property Get AnosAnt() As Integer
    AnosAnt = mProps.AnosAnt
End Property

Public Property Let AnosAnt2(ByVal value As Integer)
    mProps.AnosAnt2 = value
    mbooModificado = True
End Property

Public Property Get AnosAnt2() As Integer
    AnosAnt2 = mProps.AnosAnt2
End Property

Public Property Let AnosAnt3(ByVal value As Integer)
    mProps.AnosAnt3 = value
    mbooModificado = True
End Property

Public Property Get AnosAnt3() As Integer
    AnosAnt3 = mProps.AnosAnt3
End Property

Public Property Let ProtexionAnt(ByVal value As Boolean)
    mProps.ProtexionAnt = value
    mbooModificado = True
End Property

Public Property Get ProtexionAnt() As Boolean
    ProtexionAnt = mProps.ProtexionAnt
End Property

Public Property Let ProtexionAnt2(ByVal value As Boolean)
    mProps.ProtexionAnt2 = value
    mbooModificado = True
End Property

Public Property Get ProtexionAnt2() As Boolean
    ProtexionAnt2 = mProps.ProtexionAnt2
End Property

Public Property Let ProtexionAnt3(ByVal value As Boolean)
    mProps.ProtexionAnt3 = value
    mbooModificado = True
End Property

Public Property Get ProtexionAnt3() As Boolean
    ProtexionAnt3 = mProps.ProtexionAnt3
End Property

Public Property Let FechaActualizaCarnet(ByVal value As Date)
    mProps.FechaActualizaCarnet = value
    mbooModificado = True
End Property

Public Property Get FechaActualizaCarnet() As Date
    FechaActualizaCarnet = mProps.FechaActualizaCarnet
End Property

Public Property Let CarnetDiscapActualiza(ByVal value As Integer)
    mProps.CarnetDiscapActualiza = value
    mbooModificado = True
End Property

Public Property Get CarnetDiscapActualiza() As Integer
    CarnetDiscapActualiza = mProps.CarnetDiscapActualiza
End Property

Public Property Let TieneCarnetDiscap(ByVal value As Integer)
    mProps.TieneCarnetDiscap = value
    mbooModificado = True
End Property

Public Property Get TieneCarnetDiscap() As Integer
    TieneCarnetDiscap = mProps.TieneCarnetDiscap
End Property

Public Property Let TipoDiscapacidad(ByVal value As Integer)
    mProps.TipoDiscapacidad = value
    mbooModificado = True
End Property

Public Property Get TipoDiscapacidad() As Integer
    TipoDiscapacidad = mProps.TipoDiscapacidad
End Property


