VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RolDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type T_PROPS
    Id As Long
    transid As Long
    IdEmpleado As Long
    CodEmpleado As String
    IdGrupo1 As Long
    CodGrupo1 As String
    IdGrupo2 As Long
    CodGrupo2 As String
    IdGrupo3 As Long
    CodGrupo3 As String
    IdGrupo4 As Long
    CodGrupo4 As String
    idElemento As Long
    Codelemento As String
    AfectaSaldoEmp As Boolean
    Orden As Integer
    Valor As Currency
    numCheque As String
    banco As String
    FechaGrabado As Date
    EsNuevo As Boolean
    IdCanton As Long
    CodCanton As String
    Modificado As Boolean
    auxDebe As Long
    auxHaber As Long
    auxIdCuentaGeneral As Long
    auxIdCuentaGeneralPre As Long 'para presupuestos
    BandPagoProv As Boolean
    Tipo As Integer 'que tipo de elemento es
    Imprimir As Boolean '
    bandRenuncia As Boolean '
    BandIngreso As Boolean '
    bandEnfermedad As Boolean '
    IdTipoRol As Long
    IdTipoRol1 As Long
    IdTipoRol2 As Long
    IdTipoRol3 As Long
    IdTipoRol4 As Long
    IdTipoRol5 As Long
    IdTipoRol6 As Long
    IdTipoRol7 As Long
    IdTipoRol8 As Long
    IdTipoRol9 As Long
    IdTipoRol10 As Long
    VGP As Currency  'Variacion Gastos personales
    objGNComprobante As GNComprobante
End Type
Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Friend Property Set GNComprobante(value As GNComprobante)
    Set mProps.objGNComprobante = value
End Property

Public Property Get GNComprobante() As GNComprobante
    Set GNComprobante = mProps.objGNComprobante
End Property

Public Property Let Id(ByVal value As Long)
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "RolDetalle.Id", MSGERR_INVALIDO
    mProps.Id = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get Id() As Long
    Id = mProps.Id
End Property

Public Property Let transid(ByVal value As Long)
    mProps.transid = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get transid() As Long
    transid = mProps.transid
End Property

Public Property Let IdEmpleado(ByVal value As String)
    mProps.IdEmpleado = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get IdEmpleado() As String
    IdEmpleado = mProps.IdEmpleado
End Property

Public Property Let IdDepartamento(ByVal value As Long)
    mProps.IdGrupo1 = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get IdDepartamento() As Long
    IdDepartamento = mProps.IdGrupo1
End Property

Public Property Let idElemento(ByVal value As Long)
    mProps.idElemento = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get idElemento() As Long
    idElemento = mProps.idElemento
End Property

Public Property Get EsNuevo() As Boolean
    EsNuevo = mProps.EsNuevo
End Property
Public Property Let EsNuevo(ByVal value As Boolean) 'AGREGADO AUC 07/06/2005
    mProps.EsNuevo = value
    mbooEsNuevo = value
End Property
Public Property Let Orden(ByVal value As Integer)
      mProps.Orden = value
      Me.GNComprobante.SetModificado
End Property

Public Property Get Orden() As Integer
        Orden = mProps.Orden
End Property

Public Property Let Valor(ByVal value As Currency)
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "RolDetalle.Valor", MSGERR_INVALIDO
    mProps.Valor = value

    Me.GNComprobante.SetModificado
End Property

Public Property Get Valor() As Currency
    Valor = mProps.Valor
End Property

Public Property Let numCheque(ByVal value As String) '09/11/2000  Angel P.
    If Len(value) > 20 Then Err.Raise ERR_INVALIDO, "RolDetalle.NumCheque", MSGERR_INVALIDO
    mProps.numCheque = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get numCheque() As String            '09/11/2000  Angel P.
    numCheque = mProps.numCheque
End Property

Public Property Let banco(ByVal value As String)     '09/11/2000  Angel P.
    If Len(value) > 20 Then Err.Raise ERR_INVALIDO, "RolDetalle.Banco", MSGERR_INVALIDO
    mProps.banco = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get banco() As String                '09/11/2000  Angel P.
    banco = mProps.banco
End Property

Public Property Let FechaGrabado(ByVal value As Date)
    If Not IsDate(value) Then Err.Raise ERR_INVALIDO, "RolDetalle.FechaGrabado", MSGERR_INVALIDO
    mProps.FechaGrabado = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get FechaGrabado() As Date
    FechaGrabado = mProps.FechaGrabado
End Property

Private Sub Class_Initialize()
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objGNComprobante = Nothing
End Sub

Public Sub Recuperar(ByVal rs As Recordset)
    With rs
        mProps.Id = !Id
        mProps.transid = !transid
        If Not IsNull(!IdEmpleado) Then mProps.IdEmpleado = !IdEmpleado
        If Not IsNull(!CodEmpleado) Then mProps.CodEmpleado = !CodEmpleado
        If Not IsNull(!IdGrupo1) Then mProps.IdGrupo1 = !IdGrupo1
        If Not IsNull(!CodGrupo1) Then mProps.CodGrupo1 = !CodGrupo1
        If Not IsNull(!IdGrupo2) Then mProps.IdGrupo2 = !IdGrupo2
        If Not IsNull(!CodGrupo2) Then mProps.CodGrupo2 = !CodGrupo2
        If Not IsNull(!IdGrupo3) Then mProps.IdGrupo3 = !IdGrupo3
        If Not IsNull(!CodGrupo3) Then mProps.CodGrupo3 = !CodGrupo3
        If Not IsNull(!IdGrupo4) Then mProps.IdGrupo4 = !IdGrupo4
        If Not IsNull(!CodGrupo4) Then mProps.CodGrupo4 = !CodGrupo4
        If Not IsNull(!idElemento) Then mProps.idElemento = !idElemento
        If Not IsNull(!Codelemento) Then mProps.Codelemento = !Codelemento
        If Not IsNull(!Valor) Then mProps.Valor = !Valor
        If Not IsNull(!Orden) Then mProps.Orden = !Orden
'        If Not IsNull(!IdCanton) Then mprops.IdCanton = !IdCanton
        'If Not IsNull(!CodCanton) Then mprops.CodCanton = !CodCanton
        If Not IsNull(!BandPagoProv) Then mProps.BandPagoProv = !BandPagoProv
        If Not IsNull(!Debe) Then mProps.auxDebe = !Debe
        If Not IsNull(!Haber) Then mProps.auxHaber = !Haber
        If Not IsNull(!AfectaSaldoEmp) Then mProps.AfectaSaldoEmp = !AfectaSaldoEmp
        If Not IsNull(!Tipo) Then mProps.Tipo = !Tipo
        If Not IsNull(!Imprimir) Then mProps.Imprimir = !Imprimir
        If Not IsNull(!bandRenuncia) Then mProps.bandRenuncia = !bandRenuncia
        If Not IsNull(!bandEnfermedad) Then mProps.bandEnfermedad = !bandEnfermedad
        If Not IsNull(!BandIngreso) Then mProps.BandIngreso = !BandIngreso
        If Not IsNull(!IdTipoRol) Then mProps.IdTipoRol = !IdTipoRol
        If Not IsNull(!IdTipoRol1) Then mProps.IdTipoRol1 = !IdTipoRol1
        If Not IsNull(!IdTipoRol2) Then mProps.IdTipoRol2 = !IdTipoRol2
        If Not IsNull(!IdTipoRol3) Then mProps.IdTipoRol3 = !IdTipoRol3
        If Not IsNull(!IdTipoRol4) Then mProps.IdTipoRol4 = !IdTipoRol4
        If Not IsNull(!IdTipoRol5) Then mProps.IdTipoRol5 = !IdTipoRol5
        If Not IsNull(!IdTipoRol6) Then mProps.IdTipoRol6 = !IdTipoRol6
        If Not IsNull(!IdTipoRol7) Then mProps.IdTipoRol7 = !IdTipoRol7
        If Not IsNull(!IdTipoRol8) Then mProps.IdTipoRol8 = !IdTipoRol8
        If Not IsNull(!IdTipoRol9) Then mProps.IdTipoRol9 = !IdTipoRol9
        If Not IsNull(!IdTipoRol10) Then mProps.IdTipoRol10 = !IdTipoRol10
        
'        If Not IsNull(!auxIdCuentaGeneral) Then mProps.auxIdCuentaGeneral = !auxIdCuentaGeneral
'        If Not IsNull(!auxIdCuentaGeneralPre) Then mProps.auxIdCuentaGeneralPre = !auxIdCuentaGeneralPre

        
        mProps.FechaGrabado = !FechaGrabado
        
'        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Grabar()
    Dim sql As String, rs As Recordset
    Dim idDep As Long
    Dim idSec As Long
    idSec = Me.GNComprobante.Empresa.GNOpcion.ObtenerValor("seccion") + 1
    idDep = Me.GNComprobante.Empresa.GNOpcion.ObtenerValor("Departamento") + 1
    On Error GoTo Etiq
        sql = "SELECT * FROM roldetalle WHERE 1=0"
        Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !transid = mProps.objGNComprobante.transid
        !IdEmpleado = Me.IdEmpleado
        Select Case idDep
            Case 1: !IdGrupo1 = Me.IdGrupo1
            Case 2: !IdGrupo2 = Me.IdGrupo2
            Case 3: !IdGrupo3 = Me.IdGrupo3
            Case 4: !IdGrupo4 = Me.IdGrupo4
        End Select
        Select Case idSec
            Case 1: !IdGrupo1 = Me.IdGrupo1
            Case 2: !IdGrupo2 = Me.IdGrupo2
            Case 3: !IdGrupo3 = Me.IdGrupo3
            Case 4: !IdGrupo4 = Me.IdGrupo4
        End Select
        !idElemento = Me.idElemento
        !Orden = Me.Orden ' "" Averiguar porque no tiene valor 17/01/2001
        !Valor = Me.Valor
        '!IdCanton = Me.IdCanton
        !FechaGrabado = Now
        !BandPagoProv = 0
        !bandRenuncia = Me.bandRenuncia
        !bandEnfermedad = Me.bandEnfermedad
        !BandIngreso = Me.BandIngreso
        .Update
        If Me.VGP > 0 Then
            ActualizaEmpleado
        End If
        #If DAOLIB Then
            .Bookmark = .LastModified
        #End If
        Me.Id = !Id
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub
Private Sub ActualizaEmpleado()
Dim sql As String
    sql = "Update personal set VGP = " & Me.VGP
    sql = sql & " Where idempleado  = " & Me.IdEmpleado
    Me.GNComprobante.Empresa.EjecutarSQL sql, 1
End Sub

Public Sub GrabarxEmpleado(ByVal CodEmpleado As String)
    Dim sql As String, rs As Recordset
    Dim idDep As Long
    idDep = Me.GNComprobante.Empresa.GNOpcion.ObtenerValor("Departamento") + 1
    On Error GoTo Etiq
        sql = "SELECT * FROM roldetalle rd inner join Empleado pc on pc.idprovcli = rd.idempleado"
        sql = sql & " AND pc.codprovcli = '" & CodEmpleado & "'"
        Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !transid = mProps.objGNComprobante.transid
        !IdEmpleado = Me.IdEmpleado
        Select Case idDep
            Case 1: !IdGrupo1 = Me.IdGrupo1
            Case 2: !IdGrupo2 = Me.IdGrupo2
            Case 3: !IdGrupo3 = Me.IdGrupo3
            Case 4: !IdGrupo4 = Me.IdGrupo4
        End Select
        !idElemento = Me.idElemento
        !Orden = Me.Orden ' "" Averiguar porque no tiene valor 17/01/2001
        !Valor = Me.Valor
        !IdCanton = Me.IdCanton
        !FechaGrabado = Now
        !BandPagoProv = 0
        !bandRenuncia = Me.bandRenuncia
        !bandEnfermedad = Me.bandEnfermedad
        !BandIngreso = Me.BandIngreso
        .Update
        #If DAOLIB Then
            .Bookmark = .LastModified
        #End If
        Me.Id = !Id
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
        
        sql = "DELETE FROM RolDetalle WHERE Id = " & Me.Id
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        Me.GNComprobante.Empresa.Coneccion.Execute sql
    
    mbooEsNuevo = True
    Me.GNComprobante.SetModificado
End Sub



Public Property Let IdCanton(ByVal value As Long)
    mProps.IdCanton = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get IdCanton() As Long
    IdCanton = mProps.IdCanton
End Property


Public Property Let Modificado(value As Boolean)
     mProps.Modificado = value
End Property
Public Property Get Modificado() As Boolean
     Modificado = mProps.Modificado
End Property
Public Property Get ModificadoInd() As Boolean
     ModificadoInd = mbooModificado
End Property

Public Property Let CodEmpleado(ByVal value As String)
    Dim sql As String, rs As Recordset
    If value <> mProps.CodEmpleado Then
        If Len(value) > 0 Then
            sql = "SELECT IdProvCli FROM Empleado WHERE CodProvCli='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "RolDetalle.CodEmpleado", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdEmpleado = rs!IdProvCli
            End If
            rs.Close
        Else
            IdEmpleado = 0
        End If
        mProps.CodEmpleado = value
    Me.GNComprobante.SetModificado
    End If
    Set rs = Nothing
End Property

Public Property Get CodEmpleado() As String
    CodEmpleado = Trim$(mProps.CodEmpleado)
End Property


Public Property Let Codelemento(ByVal value As String)
    Dim sql As String, rs As Recordset
    Dim idDep As Integer
    If value <> mProps.Codelemento Then
        If Len(value) > 0 Then
            sql = "SELECT IdElemento from elemento WHERE codelemento= '" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "RolDetalle.CodElemento", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                idElemento = rs!idElemento
            End If
            rs.Close
        Else
            idElemento = 0
        End If
        mProps.Codelemento = value
    Me.GNComprobante.SetModificado
    End If
    Set rs = Nothing
End Property

Public Property Get Codelemento() As String
    Codelemento = Trim$(mProps.Codelemento)
End Property

Public Property Let CodCanton(ByVal value As String)
    Dim sql As String, rs As Recordset
    Dim idDep As Integer
    If value <> mProps.CodCanton Then
        If Len(value) > 0 Then
            sql = "SELECT Idcanton from pccanton WHERE codcanton= '" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "RolDetalle.Codcanton", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCanton = rs!IdCanton
            End If
            rs.Close
        Else
            IdCanton = 0
        End If
        mProps.CodCanton = value
    Me.GNComprobante.SetModificado
    End If
    Set rs = Nothing
End Property

Public Property Get CodCanton() As String
    CodCanton = Trim$(mProps.CodCanton)
End Property

Public Property Get auxDebe() As Boolean
    If idElemento = 0 Then Exit Property
    auxDebe = mProps.auxDebe
End Property
Public Property Let auxDebe(ByVal value As Boolean)
    mProps.auxDebe = value
    Me.GNComprobante.SetModificado
End Property
Public Property Get auxHaber() As Boolean
    If idElemento = 0 Then Exit Property
    auxHaber = mProps.auxHaber
End Property
Public Property Let auxHaber(ByVal value As Boolean)
    mProps.auxHaber = value
    Me.GNComprobante.SetModificado
End Property
Public Property Let IdGrupo1(ByVal value As Long)
    mProps.IdGrupo1 = value
    Me.GNComprobante.SetModificado
End Property
Public Property Get IdGrupo1() As Long
    IdGrupo1 = mProps.IdGrupo1
End Property

Public Property Let IdGrupo2(ByVal value As Long)
    mProps.IdGrupo2 = value
    Me.GNComprobante.SetModificado
End Property
Public Property Get IdGrupo2() As Long
    IdGrupo2 = mProps.IdGrupo2
End Property

Public Property Let IdGrupo3(ByVal value As Long)
    mProps.IdGrupo3 = value
    Me.GNComprobante.SetModificado
End Property
Public Property Get IdGrupo3() As Long
    IdGrupo3 = mProps.IdGrupo3
End Property

Public Property Let IdGrupo4(ByVal value As Long)
    mProps.IdGrupo4 = value
    Me.GNComprobante.SetModificado
End Property
Public Property Get IdGrupo4() As Long
    IdGrupo4 = mProps.IdGrupo4
End Property

Public Property Let BandPagoProv(ByVal value As Boolean)
    mProps.BandPagoProv = value
    Me.GNComprobante.SetModificado
End Property
Public Property Get BandPagoProv() As Boolean
    BandPagoProv = mProps.BandPagoProv
End Property

Public Property Let CodGrupo1(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodGrupo1 Then
        If Len(value) > 0 Then
            'Actualiza IdGrupo2 también
            sql = "SELECT IdGrupo1 FROM PCGrupo1 WHERE CodGrupo1='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Roldetalle_CodGrupo1", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdGrupo1 = rs!IdGrupo1
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdGrupo1 = 0
        End If
    End If
    mProps.CodGrupo1 = value
    Me.GNComprobante.SetModificado
End Property

Public Property Let CodGrupo2(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodGrupo2 Then
        If Len(value) > 0 Then
            'Actualiza IdGrupo2 también
            sql = "SELECT IdGrupo2 FROM PCGrupo2 WHERE CodGrupo2='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodGrupo2", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdGrupo2 = rs!IdGrupo2
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdGrupo2 = 0
        End If
    End If

    mProps.CodGrupo2 = value
    Me.GNComprobante.SetModificado
End Property

Public Property Let CodGrupo3(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodGrupo3 Then
        If Len(value) > 0 Then
            'Actualiza IdGrupo2 también
            sql = "SELECT IdGrupo3 FROM PCGrupo3 WHERE CodGrupo3='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Roldetalle_CodGrupo3", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdGrupo3 = rs!IdGrupo3
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdGrupo3 = 0
        End If
    End If

    mProps.CodGrupo3 = value
    Me.GNComprobante.SetModificado
End Property

Public Property Let CodGrupo4(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodGrupo4 Then
        If Len(value) > 0 Then
            'Actualiza IdGrupo2 también
            sql = "SELECT IdGrupo4 FROM PCGrupo4 WHERE CodGrupo4='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Roldetalle_CodGrupo4", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdGrupo4 = rs!IdGrupo4
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdGrupo4 = 0
        End If
    End If
    mProps.CodGrupo4 = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get CodGrupo1() As String
    CodGrupo1 = Trim$(mProps.CodGrupo1)
End Property

Public Property Get CodGrupo2() As String
    CodGrupo2 = Trim$(mProps.CodGrupo2)
End Property

Public Property Get CodGrupo3() As String
    CodGrupo3 = Trim$(mProps.CodGrupo3)
End Property
Public Property Get CodGrupo4() As String
    CodGrupo4 = Trim$(mProps.CodGrupo4)
End Property
'
Public Property Let BandAfectaSaldoEmp(ByVal value As Boolean)
    mProps.AfectaSaldoEmp = value
    Me.GNComprobante.SetModificado
End Property
Public Property Get BandAfectaSaldoEmp() As Boolean
    BandAfectaSaldoEmp = mProps.AfectaSaldoEmp
End Property
Public Property Get auxIdCuentaGeneral() As Long
    auxIdCuentaGeneral = mProps.auxIdCuentaGeneral
End Property
Public Property Let auxIdCuentaGeneral(ByVal value As Long)
    mProps.auxIdCuentaGeneral = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get Tipo() As Integer
    Tipo = mProps.Tipo
End Property
Public Property Let Tipo(ByVal value As Integer)
    mProps.Tipo = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get Imprimir() As Boolean
    Imprimir = mProps.Imprimir
End Property
Public Property Let Imprimir(ByVal value As Boolean)
    mProps.Imprimir = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get bandRenuncia() As Boolean
    bandRenuncia = mProps.bandRenuncia
End Property

Public Property Let bandRenuncia(ByVal value As Boolean)
    mProps.bandRenuncia = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get BandIngreso() As Boolean
    BandIngreso = mProps.BandIngreso
End Property

Public Property Let BandIngreso(ByVal value As Boolean)
    mProps.BandIngreso = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get bandEnfermedad() As Boolean
    bandEnfermedad = mProps.bandEnfermedad
End Property

Public Property Let bandEnfermedad(ByVal value As Boolean)
    mProps.bandEnfermedad = value
    Me.GNComprobante.SetModificado
End Property
'AUC Para tomar el id del presupuesto desde el objeto rol
Public Property Get auxIdCuentaGeneralPre() As Long
    auxIdCuentaGeneralPre = mProps.auxIdCuentaGeneralPre
End Property
Public Property Let auxIdCuentaGeneralPre(ByVal value As Long)
    mProps.auxIdCuentaGeneralPre = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get IdTipoRol() As Long
    IdTipoRol = mProps.IdTipoRol
End Property
Public Property Let IdTipoRol(ByVal value As Long)
    mProps.IdTipoRol = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get IdTipoRol1() As Long
    IdTipoRol1 = mProps.IdTipoRol1
End Property
Public Property Let IdTipoRol1(ByVal value As Long)
    mProps.IdTipoRol1 = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get IdTipoRol2() As Long
    IdTipoRol2 = mProps.IdTipoRol2
End Property
Public Property Let IdTipoRol2(ByVal value As Long)
    mProps.IdTipoRol2 = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get IdTipoRol3() As Long
    IdTipoRol3 = mProps.IdTipoRol3
End Property
Public Property Let IdTipoRol3(ByVal value As Long)
    mProps.IdTipoRol3 = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get IdTipoRol4() As Long
    IdTipoRol4 = mProps.IdTipoRol4
End Property
Public Property Let IdTipoRol4(ByVal value As Long)
    mProps.IdTipoRol4 = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get IdTipoRol5() As Long
    IdTipoRol5 = mProps.IdTipoRol5
End Property
Public Property Let IdTipoRol5(ByVal value As Long)
    mProps.IdTipoRol5 = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get IdTipoRol6() As Long
    IdTipoRol6 = mProps.IdTipoRol6
End Property
Public Property Let IdTipoRol6(ByVal value As Long)
    mProps.IdTipoRol6 = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get IdTipoRol7() As Long
    IdTipoRol7 = mProps.IdTipoRol7
End Property
Public Property Let IdTipoRol7(ByVal value As Long)
    mProps.IdTipoRol7 = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get IdTipoRol8() As Long
    IdTipoRol8 = mProps.IdTipoRol8
End Property
Public Property Let IdTipoRol8(ByVal value As Long)
    mProps.IdTipoRol8 = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get IdTipoRol9() As Long
    IdTipoRol9 = mProps.IdTipoRol9
End Property
Public Property Let IdTipoRol9(ByVal value As Long)
    mProps.IdTipoRol9 = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get IdTipoRol10() As Long
    IdTipoRol10 = mProps.IdTipoRol10
End Property
Public Property Let IdTipoRol10(ByVal value As Long)
    mProps.IdTipoRol10 = value
    Me.GNComprobante.SetModificado
End Property

Public Property Get VGP() As Currency
    VGP = mProps.VGP
End Property
Public Property Let VGP(ByVal value As Currency)
    mProps.VGP = value
    Me.GNComprobante.SetModificado
End Property



