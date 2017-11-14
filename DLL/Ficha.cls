VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ficha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    idficha As Long
    Historia As String 'no vale
    HistoriaClinica As Long 'reemplaza a historia
    IdCliente As Long
    Sexo As Byte
    Cedula As String
    idCentro As Long
    Piso As Integer
    Nombre As String
    FechaIngreso As Date
    Hora As Date
    Edad As Integer
    Procedencia As String
    IdCiudad As Long
    Direccion As String
    Telefono As String
    Ocupacion As String
    Religion As String
    AntecedentesP As String
    AntecedentesF As String
    
    
    CodCliente As String
    CodCiudad As String
    
    CodAntePersonal As String
    DescAntePersonal As String
    Familiar As String
    CodAnteFamiliar As String
    DescAnteFamiliar As String
    FechaNacimiento As Date
    Etnia As String
    Discapacidad As String
    Nota As String
    IdTipoDocumento As Integer
    codtipoDocumento As String
    objEmpresa As Empresa
    
    
End Type
Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean
Event Refresh(Tipo As Integer)

Public Property Let idficha(ByVal value As Long)
    mProps.idficha = value
    mbooModificado = True
End Property

Public Property Get idficha() As Long
    idficha = mProps.idficha
End Property

Friend Property Set objEmpresa(ByVal obj As Empresa)
    Set mProps.objEmpresa = obj
End Property

Public Property Get objEmpresa() As Empresa
    Set objEmpresa = mProps.objEmpresa
End Property

Public Property Let Sexo(ByVal value As Byte)
    '0=Femenino, 1=Masculino
    If (Not IsNumeric(value)) Or (Val(value) < 0 Or Val(value) > 1) Then Err.Raise ERR_INVALIDO, "PCEmpleado.Sexo", MSGERR_INVALIDO
    mProps.Sexo = value
    mbooModificado = True
End Property

Public Property Get Sexo() As Byte
    Sexo = mProps.Sexo
End Property

Public Property Let IdCliente(ByVal value As Long)
    mProps.IdCliente = value
    mbooModificado = True
End Property

Public Property Get IdCliente() As Long
    IdCliente = mProps.IdCliente
End Property

Public Property Let Cedula(ByVal value As String)
    If Len(value) > 10 Then Err.Raise ERR_INVALIDO, "Ficha.Cedula", MSGERR_INVALIDO
    mProps.Cedula = value
    mbooModificado = True
End Property

Public Property Get Cedula() As String
    Cedula = mProps.Cedula
End Property

Public Property Let idCentro(ByVal value As Long)
    mProps.idCentro = value
    mbooModificado = True
End Property

Public Property Get idCentro() As Long
    idCentro = mProps.idCentro
End Property

Public Property Get Piso() As Integer
    Piso = mProps.Piso
End Property
Public Property Let Piso(ByVal value As Integer)
    mProps.Piso = value
    mbooModificado = True
End Property

Public Property Let Nombre(ByVal value As String)
    mProps.Nombre = value
    mbooModificado = True
End Property

Public Property Get Nombre() As String
    Nombre = mProps.Nombre
End Property

Public Property Let FechaIngreso(ByVal value As Date)
    If Not IsDate(value) Then Err.Raise ERR_INVALIDO, "Ficha.FechaIngreso", MSGERR_INVALIDO
    mProps.FechaIngreso = value
    mbooModificado = True
End Property

Public Property Get FechaIngreso() As Date
    FechaIngreso = mProps.FechaIngreso
End Property

Public Property Let Hora(ByVal value As Date)
    If Not IsDate(value) Then Err.Raise ERR_INVALIDO, "Ficha.hora", MSGERR_INVALIDO
    mProps.Hora = value
    mbooModificado = True
End Property

Public Property Get Hora() As Date
    Hora = mProps.Hora
End Property

Public Property Get Edad() As Integer
    Edad = mProps.Edad
End Property
Public Property Let Edad(ByVal value As Integer)
    mProps.Edad = value
    mbooModificado = True
End Property

Public Property Let Procedencia(ByVal value As String)
    mProps.Procedencia = value
    mbooModificado = True
End Property

Public Property Get Procedencia() As String
    Procedencia = mProps.Procedencia
End Property
Public Property Let IdCiudad(ByVal value As Long)
    mProps.IdCiudad = value
    mbooModificado = True
End Property

Public Property Get IdCiudad() As Long
    IdCiudad = mProps.IdCiudad
End Property

Public Property Let Direccion(ByVal value As String)
    mProps.Direccion = value
    mbooModificado = True
End Property

Public Property Get Direccion() As String
    Direccion = mProps.Direccion
End Property

Public Property Let Telefono(ByVal value As String)
    mProps.Telefono = value
    mbooModificado = True
End Property

Public Property Get Telefono() As String
    Telefono = mProps.Telefono
End Property

Public Property Let Ocupacion(ByVal value As String)
    mProps.Ocupacion = value
    mbooModificado = True
End Property

Public Property Get Ocupacion() As String
    Ocupacion = mProps.Ocupacion
End Property

Public Property Let Religion(ByVal value As String)
    mProps.Religion = value
    mbooModificado = True
End Property

Public Property Get Religion() As String
    Religion = mProps.Religion
End Property

Public Property Let idmedico(ByVal value As Long)
'    mProps.idmedico = Value
'    mbooModificado = True
End Property

Public Property Get idmedico() As Long
    'idmedico = mProps.idmedico
End Property

Public Property Let Sintomas(ByVal value As String)
'    mProps.Sintomas = Value
'    mbooModificado = True
End Property

Public Property Get Sintomas() As String
'    Sintomas = mProps.Sintomas
End Property

Public Property Let ExamenFisico(ByVal value As String)
    'mProps.ExamenFisico = Value
    'mbooModificado = True
End Property

Public Property Get ExamenFisico() As String
    'ExamenFisico = mProps.ExamenFisico
End Property

Public Property Let DX(ByVal value As String)
'    mProps.DX = Value
'    mbooModificado = True
End Property

Public Property Get DX() As String
'    DX = mProps.DX
End Property

Public Property Let Antesedentes(ByVal value As String)
'    mProps.Antesedentes = Value
'    mbooModificado = True
End Property

Public Property Get Antesedentes() As String
'    Antesedentes = mProps.Antesedentes
End Property

Public Property Let Peso(ByVal value As String)
'    mProps.Peso = Value
'    mbooModificado = True
End Property

Public Property Get Peso() As String
'    Peso = mProps.Peso
End Property

Public Property Let TipoSangre(ByVal value As String)
'    mProps.TipoSangre = Value
'    mbooModificado = True
End Property

Public Property Get TipoSangre() As String
'    TipoSangre = mProps.TipoSangre
End Property

Public Property Let Infancia(ByVal value As String)
'    mProps.Infancia = Value
'    mbooModificado = True
End Property

Public Property Get Infancia() As String
'    Infancia = mProps.Infancia
End Property

Public Property Let Pubertad(ByVal value As String)
'    mProps.Pubertad = Value
'    mbooModificado = True
End Property

Public Property Get Pubertad() As String
'    Pubertad = mProps.Pubertad
End Property

Public Property Let Adulto(ByVal value As String)
'    mProps.Adulto = Value
'    mbooModificado = True
End Property

Public Property Get Adulto() As String
'    Adulto = mProps.Adulto
End Property

Public Property Let Habito(ByVal value As String)
'    mProps.Habito = Value
'    mbooModificado = True
End Property

Public Property Get Habito() As String
'    Habito = mProps.Habito
End Property
Public Property Let Alergia(ByVal value As String)
'    mProps.Alergia = Value
'    mbooModificado = True
End Property

Public Property Get Alergia() As String
'    Alergia = mProps.Alergia
End Property

Public Property Let Medicamentos(ByVal value As String)
'    mProps.Medicamentos = Value
'    mbooModificado = True
End Property

Public Property Get Medicamentos() As String
'    Medicamentos = mProps.Medicamentos
End Property

Public Property Let Corticostiroides(ByVal value As String)
'    mProps.Corticostiroides = Value
'    mbooModificado = True
End Property

Public Property Get Corticostiroides() As String
'    Corticostiroides = mProps.Corticostiroides
End Property

Public Property Let DerivadosFE(ByVal value As String)
'    mProps.DerivadosFE = Value
'    mbooModificado = True
End Property

Public Property Get DerivadosFE() As String
'    DerivadosFE = mProps.DerivadosFE
End Property

Public Property Let DerivadosRA(ByVal value As String)
'    mProps.DerivadosRA = Value
'    mbooModificado = True
End Property

Public Property Get DerivadosRA() As String
'    DerivadosRA = mProps.DerivadosRA
End Property

Public Property Let Antidiabeticos(ByVal value As String)
'    mProps.Antidiabeticos = Value
'    mbooModificado = True
End Property

Public Property Get Antidiabeticos() As String
'    Antidiabeticos = mProps.Antidiabeticos
End Property

Public Property Let CP(ByVal value As String)
'    mProps.CP = Value
'    mbooModificado = True
End Property

Public Property Get CP() As String
    'CP = mProps.CP
End Property

Public Property Let Pulso(ByVal value As String)
'    mProps.Pulso = Value
'    mbooModificado = True
End Property

Public Property Get Pulso() As String
'    Pulso = mProps.Pulso
End Property



Public Property Let Antibioticos(ByVal value As String)
'    mProps.Antibioticos = Value
'    mbooModificado = True
End Property

Public Property Get Antibioticos() As String
'    Antibioticos = mProps.Antibioticos
End Property


Public Property Let Temperatura(ByVal value As String)
'    mProps.Temperatura = Value
'    mbooModificado = True
End Property

Public Property Get Temperatura() As String
'    Temperatura = mProps.Temperatura
End Property

Public Property Let Respiracion(ByVal value As String)
'    mProps.Respiracion = Value
'    mbooModificado = True
End Property

Public Property Get Respiracion() As String
'    Respiracion = mProps.Respiracion
End Property

Public Property Let Pupilas(ByVal value As String)
'    mProps.Pupilas = Value
'    mbooModificado = True
End Property

Public Property Get Pupilas() As String
'    Pupilas = mProps.Pupilas
End Property

Public Property Let Otros(ByVal value As String)
'    mProps.Otros = Value
'    mbooModificado = True
End Property

Public Property Get Otros() As String
'    Otros = mProps.Otros
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
    Set mProps.objEmpresa = Nothing
End Sub

Public Sub Recuperar(ByVal id As Long)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT f.*,pc.codprovcli,pcc.codCanton as codciudad,a.codtipoDocumento FROM Ficha f left join pcprovcli pc on pc.idprovcli=f.idcliente "
    sql = sql & " Left Join pccanton pcc on pcc.idcanton = f.idciudad"
    sql = sql & " Left Join Anexo_TipoDocumento a on a.id = f.idTipoDocumento"
    sql = sql & " Where IdFicha = " & id
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
        
    With rs
        Do While Not .EOF
            mProps.idficha = !idficha
            If Not IsNull(!Historia) Then mProps.HistoriaClinica = !Historia
            If Not IsNull(!IdCliente) Then mProps.IdCliente = !IdCliente
            If Not IsNull(!CodProvcli) Then mProps.CodCliente = !CodProvcli
            If Not IsNull(!Sexo) Then mProps.Sexo = !Sexo
            If Not IsNull(!Cedula) Then mProps.Cedula = !Cedula

            If Not IsNull(!Nombre) Then mProps.Nombre = !Nombre
            If Not IsNull(!Procedencia) Then mProps.Procedencia = !Procedencia
            If Not IsNull(!IdCiudad) Then mProps.IdCiudad = !IdCiudad
            If Not IsNull(!CodCiudad) Then mProps.CodCiudad = !CodCiudad
            If Not IsNull(!Direccion) Then mProps.Direccion = !Direccion
            If Not IsNull(!Telefono) Then mProps.Telefono = !Telefono
            If Not IsNull(!Ocupacion) Then mProps.Ocupacion = !Ocupacion
            If Not IsNull(!Religion) Then mProps.Religion = !Religion
            If Not IsNull(!Familiar) Then mProps.Familiar = !Familiar
            If Not IsNull(!AntecedentesP) Then mProps.AntecedentesP = !AntecedentesP
            If Not IsNull(!AntecedentesF) Then mProps.AntecedentesF = !AntecedentesF
            
            If Not IsNull(!FechaNacimiento) Then mProps.FechaNacimiento = !FechaNacimiento
            If Not IsNull(!Etnia) Then mProps.Etnia = !Etnia
            If Not IsNull(!Discapacidad) Then mProps.Discapacidad = !Discapacidad
            If Not IsNull(!Nota) Then mProps.Nota = !Nota
            mProps.IdTipoDocumento = !IdTipoDocumento
            If Not IsNull(!codtipoDocumento) Then mProps.codtipoDocumento = !codtipoDocumento 'Agregado jeaa 15/05/2007
            
            rs.MoveNext
        Loop
            .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Grabar(ByVal idficha As Long)
    Dim sql As String, rs As Recordset
    On Error GoTo Etiq
            
    If EsNuevo Then
        mProps.HistoriaClinica = AsignaNumFicha
    End If
        sql = "SELECT * FROM Ficha WHERE idFicha= " & idficha
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            sql = "SELECT * FROM Ficha WHERE 1=0"
            Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
            rs.AddNew
        Else
        End If
    With rs
            !Historia = mProps.HistoriaClinica
            !IdCliente = mProps.IdCliente
            !Sexo = mProps.Sexo
            !Cedula = mProps.Cedula
            !Nombre = mProps.Nombre
            !Procedencia = mProps.Procedencia
            !IdCiudad = mProps.IdCiudad
            !Direccion = mProps.Direccion
            !Telefono = mProps.Telefono
            !Ocupacion = mProps.Ocupacion
            !Religion = mProps.Religion
            !Familiar = mProps.Familiar
            !AntecedentesP = mProps.AntecedentesP
            !AntecedentesF = mProps.AntecedentesF
            !FechaNacimiento = mProps.FechaNacimiento
            !Etnia = mProps.Etnia
            !Discapacidad = mProps.Discapacidad
            !Nota = mProps.Nota
            !IdTipoDocumento = mProps.IdTipoDocumento
            .Update
            .Move 0
            mProps.idficha = !idficha
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    Me.objEmpresa.RaiseRefreshEvent REFRESH_FICHA
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Public Property Let Historia(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "Ficha_Historia", MSGERR_INVALIDO
    End If
    
    mProps.Historia = value
    mbooModificado = True
    
End Property

Public Property Get Historia() As String
    Historia = mProps.Historia
End Property

Public Function CreaFichaDetalle() As fichaDetalle
    Dim obj As fichaDetalle
    Set obj = New fichaDetalle
    Set obj.objFicha = Me
    Set CreaFichaDetalle = obj
    Set obj = Nothing
End Function


Public Function ListaDetalle() As Variant
    Dim sql As String, rs As Recordset
    sql = "SELECT id,Piso,codcentro,fechaIngreso,hora,f.nombre "
    sql = sql & " FROM FichaDetalle fd "
    sql = sql & " LEFT JOIN GNCentroCosto cc on cc.idcentro = fd.idcentro"
    sql = sql & " LEFT JOIN FCVendedor f on f.idvendedor = fd.idmedico"
    sql = sql & " Where IdFicha = " & mProps.idficha
    sql = sql & " ORDER BY id"
    Set rs = Me.objEmpresa.OpenRecordset(sql)
    'If BandRS Then
        Set ListaDetalle = rs
   ' Else
'        ListaDetalle = MiGetRows(rs)
'        rs.Close
    'End If
    Set rs = Nothing
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
Public Sub RaiseRefreshEvent(Tipo As Integer)
    'RaiseEvent Refres(Tipo) 'no vale
End Sub

Public Property Let CodCliente(ByVal value As String)
    Dim sql As String
    Dim rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCliente Then
        If Len(value) > 0 Then
            'Actualiza IdCentro también
            sql = "SELECT IdProvCli FROM PCProvCli WHERE CodProvCli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Ficha.CodCliente", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCliente = rs!IdProvcli
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdCliente = 0
        End If
        mProps.CodCliente = value
        mbooModificado = True
    End If
End Property

Public Property Get CodCliente() As String
    CodCliente = Trim$(mProps.CodCliente)
End Property


Public Property Let CodCiudad(ByVal value As String)
    Dim sql As String
    Dim rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCiudad Then
        If Len(value) > 0 Then
            'Actualiza IdCentro también
            sql = "SELECT IdCanton FROM PCCanton WHERE CodCanton='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Ficha.CodCanton", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCiudad = rs!IdCanton
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdCiudad = 0
        End If
        mProps.CodCiudad = value
        mbooModificado = True
    End If
End Property

Public Property Get CodCiudad() As String
    CodCiudad = Trim$(mProps.CodCiudad)
End Property

Private Function AsignaNumFicha() As Long
    Dim num As Long, sql As String, rs As Recordset, rs2 As Recordset
    
    num = Val(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("FichaNumHisSig"))

    'Si es que ya está asignado un número por UI
        If mProps.HistoriaClinica <> 0 Then num = mProps.HistoriaClinica   'Coge ese número
    
    
    'Asegura que no esté ocupado el número
    Do
        
            sql = "SELECT idFicha FROM Ficha " & _
                  "WHERE  Historia =" & num
        
        Set rs2 = mProps.objEmpresa.OpenRecordset(sql)
        If rs2.EOF Then Exit Do
        num = num + 1
        rs2.Close
    Loop
    rs2.Close
    
        'Actualiza el numero siguiente
        gobjMain.EmpresaActual.GNOpcion.AsignarValor ("FichaNumHisSig"), num + 1
        gobjMain.EmpresaActual.GNOpcion.GrabarGNOpcion2
        
    Set rs2 = Nothing
    AsignaNumFicha = num
End Function

Public Property Let HistoriaClinica(ByVal value As Long)
    mProps.HistoriaClinica = value
    mbooModificado = True
End Property

Public Property Get HistoriaClinica() As Long
    HistoriaClinica = mProps.HistoriaClinica
End Property

Public Property Let Familiar(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Ficha_Familiar", MSGERR_INVALIDO
    End If
    
    mProps.Familiar = value
    mbooModificado = True
    
End Property

Public Property Get Familiar() As String
    Familiar = mProps.Familiar
End Property
Public Sub grabarAntePersonal(ByVal IdAnte As Long)
    Dim sql As String, rs As Recordset
    On Error GoTo Etiq
        
    sql = "SELECT * FROM FAntecedentePersonal WHERE idAntecedente= " & IdAnte
   Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
       
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sql = "SELECT * FROM FAntecedentePersonal WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    Else
    End If
    With rs
            !idficha = mProps.idficha
            !Codigo = mProps.CodAntePersonal
            !Descripcion = mProps.DescAntePersonal
        .Update
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    'Me.objEmpresa.RaiseRefreshEvent REFRESH_FICHA
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub
Public Property Let CodAntePersonal(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "Ficha_CodAntePersonal", MSGERR_INVALIDO
    End If
    mProps.CodAntePersonal = value
    mbooModificado = True
End Property

Public Property Get CodAntePersonal() As String
    CodAntePersonal = mProps.CodAntePersonal
End Property

Public Property Let DescAntePersonal(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Ficha_CodAntePersonal", MSGERR_INVALIDO
    End If
    mProps.DescAntePersonal = value
    mbooModificado = True
End Property

Public Property Get DescAntePersonal() As String
    DescAntePersonal = mProps.DescAntePersonal
End Property
'---------------------
Public Sub grabarAnteFamiliar(ByVal IdAnte As Long)
    Dim sql As String, rs As Recordset
    On Error GoTo Etiq
        
    sql = "SELECT * FROM FAntecedenteFamiliar WHERE idAntecedenteF= " & IdAnte
   Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
       
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sql = "SELECT * FROM FAntecedenteFamiliar WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    Else
    End If
    With rs
            !idficha = mProps.idficha
            !Codigo = mProps.CodAnteFamiliar
            !Descripcion = mProps.DescAnteFamiliar
        .Update
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    'Me.objEmpresa.RaiseRefreshEvent REFRESH_FICHA
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub
Public Property Let CodAnteFamiliar(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "Ficha_CodAnteFamiliar", MSGERR_INVALIDO
    End If
    mProps.CodAnteFamiliar = value
    mbooModificado = True
End Property

Public Property Get CodAnteFamiliar() As String
    CodAnteFamiliar = mProps.CodAnteFamiliar
End Property

Public Property Let DescAnteFamiliar(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Ficha_CodAnteFamiliar", MSGERR_INVALIDO
    End If
    mProps.DescAnteFamiliar = value
    mbooModificado = True
End Property

Public Property Get DescAnteFamiliar() As String
    DescAnteFamiliar = mProps.DescAnteFamiliar
End Property
Public Sub BorradoFisico(cod As Long)
    

End Sub


Public Function EliminaDetalle(cod, idficha) As fichaDetalle
    Dim nuevo As fichaDetalle
    Set nuevo = New fichaDetalle
    Set nuevo.objFicha = Me
    nuevo.BorradoFisico Val(cod), Val(idficha)
    Set EliminaDetalle = nuevo
    Set nuevo = Nothing
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

Public Property Let AntecedenteP(ByVal value As String)
    If Len(value) > 512 Then Err.Raise ERR_INVALIDO, "Ficha.Antecedetep", MSGERR_INVALIDO
    mProps.AntecedentesP = value
    mbooModificado = True
End Property

Public Property Get AntecedenteP() As String
    AntecedenteP = mProps.AntecedentesP
End Property


Public Property Let AntecedenteF(ByVal value As String)
    If Len(value) > 512 Then Err.Raise ERR_INVALIDO, "Ficha.AntecedeteF", MSGERR_INVALIDO
    mProps.AntecedentesF = value
    mbooModificado = True
End Property

Public Property Get AntecedenteF() As String
    AntecedenteF = mProps.AntecedentesF
End Property
Public Function VerificaCedula(ByVal NumCedula As String) As Boolean
        VerificaCedula = VerificaDocumento(NumCedula)
End Function

Public Function VerificaCedulaRepetido() As Boolean
    Dim rs As Recordset, sql As String
    VerificaCedulaRepetido = False
    sql = "SELECT  Cedula FROM Ficha"
    sql = sql & " WHERE (Cedula='" & mProps.Cedula & "') AND "
    sql = sql & " (IdFicha<>" & mProps.idficha & ")"
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        VerificaCedulaRepetido = True
    End If
    rs.Close
End Function

Public Sub RecuperarXTrans(ByVal cod As Long)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT ccd.descripciondet,f.* from ficha f inner join fichadetalle fd "
    sql = sql & " left join gncentrocostodetalle ccd on ccd.idfichadet = fd.id on fd.idficha=f.idficha "
    sql = sql & " Where ccd.idcentrodet = " & cod
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
        
    With rs
        Do While Not .EOF
            mProps.idficha = !idficha
            If Not IsNull(!Historia) Then mProps.HistoriaClinica = !Historia
            If Not IsNull(!IdCliente) Then mProps.IdCliente = !IdCliente

            If Not IsNull(!Sexo) Then mProps.Sexo = !Sexo
            If Not IsNull(!Cedula) Then mProps.Cedula = !Cedula
            If Not IsNull(!Nombre) Then mProps.Nombre = !Nombre
            If Not IsNull(!Procedencia) Then mProps.Procedencia = !Procedencia
            If Not IsNull(!IdCiudad) Then mProps.IdCiudad = !IdCiudad
           
            If Not IsNull(!Direccion) Then mProps.Direccion = !Direccion
            If Not IsNull(!Telefono) Then mProps.Telefono = !Telefono
            If Not IsNull(!Ocupacion) Then mProps.Ocupacion = !Ocupacion
            If Not IsNull(!Religion) Then mProps.Religion = !Religion
            If Not IsNull(!Familiar) Then mProps.Familiar = !Familiar
            If Not IsNull(!AntecedentesP) Then mProps.AntecedentesP = !AntecedentesP
            If Not IsNull(!AntecedentesF) Then mProps.AntecedentesF = !AntecedentesF
            rs.MoveNext
        Loop
            .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Function RecuperaFichaDetalleXCCDetalle(key As Variant) As fichaDetalle
    Dim obj As fichaDetalle
    Set obj = New fichaDetalle
    Set obj.objFicha = Me
    On Error Resume Next
    obj.recuperarxCCDetalle key
    If Err.Number = 0 Then Set RecuperaFichaDetalleXCCDetalle = obj
    Err.Clear
    On Error GoTo 0
    Set obj = Nothing
End Function


Public Function Eliminar(cod As Long) As Boolean
    Dim sql As String
    Dim rs As Recordset
    On Error GoTo CapturaError
    'revisamos que no tenga trans relacionada
    sql = "Select gn.codtrans,gn.numtrans,gn.idcentrohijo from gncomprobante gn "
    sql = sql & " Inner Join gncentrocostodetalle ccd "
    sql = sql & "inner join FichaDetalle fd Inner join Ficha f on f.idficha = fd.idficha on fd.id = ccd.idfichadet"
    sql = sql & " on ccd.idcentrodet = gn.idcentrohijo "
    sql = sql & " where f.idficha = " & cod
    Set rs = Me.objEmpresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        MsgBox " No puede borrar la ficha tiene registros relacionados" & Chr(13) & rs!CodTrans & rs!NumTrans, vbCritical
        Eliminar = False
        Exit Function
    End If
    
    'cambia el estado de la habitación
    sql = "update cc set cc.bandocupado = 0"
    sql = sql & " from gncentrocosto cc  "
    sql = sql & " INNER JOIN FichaDetalle fd"
    sql = sql & " inner join ficha f on f.idficha= fd.idficha"
    sql = sql & "  on fd.idcentro = cc.idcentro"
    sql = sql & " Where f.idficha = " & cod
    objEmpresa.EjecutarSQL sql, 1
    
    sql = "delete from Ficha where IdFicha= " & cod
    objEmpresa.EjecutarSQL sql, 1
    Eliminar = True
    Me.objEmpresa.RaiseRefreshEvent REFRESH_FICHA
    Set rs = Nothing
    Exit Function
CapturaError:
    MsgBox Err.Description
    Set rs = Nothing
    Eliminar = False
    Exit Function
End Function

Public Property Let Etnia(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Ficha_Etnia", MSGERR_INVALIDO
    End If
    mProps.Etnia = value
    mbooModificado = True
End Property

Public Property Get Etnia() As String
    Etnia = mProps.Etnia
End Property

Public Property Let Discapacidad(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Ficha_Discapacidad", MSGERR_INVALIDO
    End If
    mProps.Discapacidad = value
    mbooModificado = True
End Property

Public Property Get Discapacidad() As String
    Discapacidad = mProps.Discapacidad
End Property

Public Property Let Nota(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Ficha_Nota", MSGERR_INVALIDO
    End If
    mProps.Nota = value
    mbooModificado = True
End Property

Public Property Get Nota() As String
    Nota = mProps.Nota
End Property

Public Property Let FechaNacimiento(ByVal value As Date)
    If Not IsDate(value) Then Err.Raise ERR_INVALIDO, "Ficha.FechaNacimiento", MSGERR_INVALIDO
    mProps.FechaNacimiento = value
    mbooModificado = True
End Property

Public Property Get FechaNacimiento() As Date
    FechaNacimiento = mProps.FechaNacimiento
End Property
Public Property Let codtipoDocumento(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.codtipoDocumento Then
        If Len(value) > 0 Then
            'Actualiza IdTipoDocumento también
            sql = "SELECT Id FROM Anexo_TipoDocumento WHERE CodTipoDocumento ='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Ficha_CodTipoDocumento", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdTipoDocumento = rs!id
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdTipoDocumento = 0
        End If
    End If

    mProps.codtipoDocumento = value
    
    mbooModificado = True
End Property

Public Property Get codtipoDocumento() As String
    codtipoDocumento = Trim$(mProps.codtipoDocumento)
End Property
Public Property Get IdTipoDocumento() As Long
    IdTipoDocumento = mProps.IdTipoDocumento
End Property

Public Property Let IdTipoDocumento(ByVal value As Long)
    mProps.IdTipoDocumento = value
    mbooModificado = True
End Property
