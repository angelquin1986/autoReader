VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Elementos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    idElemento As Long
    Codelemento As String
    Nombre As String
    Descripcion As String
    Formula As String
    Meses As String
    Tipo As Long
    Editable As Boolean
    Imprimir As Boolean
    Orden As Long
    BandActivo As Boolean
    Visible As Boolean
    fechagrabado As Date
    AsignaValorRg As Long
    ValidarAsignar As Boolean 'AUC 23/01/07
    MostrarEnReporte As Boolean 'AUC 23/02/07
    MostrarEnprovision As Boolean
    IdDepartamento As Long
    BandAcumular As Boolean 'AUC 01/40/07
    IdCuentaDebe As Long
    IdCuentaHaber As Long
    CodCuentaDebe As String
    CodCuentaHaber As String
    BandAfectaEmp As Boolean 'para saber si el total de este elemento se va o no como cuenta por pagar
    Debe As Boolean
    Haber As Boolean
    IdTipoRol As Long
    CodTipoRol As String
    IdTipoRol1 As Long
    CodTipoRol1 As String
    IdTipoRol2 As Long
    CodTipoRol2 As String
    IdTipoRol3 As Long
    CodTipoRol3 As String
    BandGastoDist As Boolean
    BandBanco As Boolean
    IdTipoRol4 As Long
    CodTipoRol4 As String
    IdTipoRol5 As Long
    CodTipoRol5 As String
    IdTipoRol6 As Long
    CodTipoRol6 As String
    IdTipoRol7 As Long
    CodTipoRol7 As String
    IdTipoRol8 As Long
    CodTipoRol8 As String
    IdTipoRol9 As Long
    CodTipoRol9 As String
    IdTipoRol10 As Long
    CodTipoRol10 As String
    BandGastoDistPorc As Boolean
    objEmpresa As Empresa
End Type
Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Friend Property Set Empresa(ByVal obj As Empresa)
    Set mProps.objEmpresa = obj
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Let IdDepartamento(ByVal value As Long)
    mProps.IdDepartamento = value
    mbooModificado = True
End Property

Public Property Get IdDepartamento() As Long
    IdDepartamento = mProps.IdDepartamento
End Property
'propiedad Agregada para permitir saber que elemento se podra llenar toda la columna en el rol general
'auc 22/06/2005
Public Property Let AsignaValorRg(ByVal value As Boolean)
'    mProps.AsignaValorRg = value
'    mbooModificado = True
End Property

Public Property Get AsignaValorRg() As Boolean
'    AsignaValorRg = mProps.AsignaValorRg
End Property

Public Property Let Nombre(ByVal value As String)
    If Len(value) > 50 Then Err.Raise ERR_INVALIDO, "Elemento.Nombre", MSGERR_INVALIDO
    mProps.Nombre = value
    mbooModificado = True
End Property

Public Property Get Nombre() As String
    Nombre = mProps.Nombre
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 80 Then Err.Raise ERR_INVALIDO, "Elemento.Descripcion", MSGERR_INVALIDO
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = mProps.Descripcion
End Property

Public Property Let Formula(ByVal value As String)
    mProps.Formula = value
    mbooModificado = True
End Property

Public Property Get Formula() As String
    Formula = mProps.Formula
End Property

Public Property Let Meses(ByVal value As String)
    If Len(value) > 30 Then Err.Raise ERR_INVALIDO, "Elemento.Meses", MSGERR_INVALIDO
    mProps.Meses = value
    mbooModificado = True
End Property

Public Property Get Meses() As String
    Meses = mProps.Meses
End Property

Public Property Let Tipo(ByVal value As Long)
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "Elemento.Tipo", MSGERR_INVALIDO
    mProps.Tipo = value
    mbooModificado = True
End Property

Public Property Get Tipo() As Long
    Tipo = mProps.Tipo
End Property

Public Property Let Editable(ByVal value As Boolean)
    mProps.Editable = value
    mbooModificado = True
End Property

Public Property Get Editable() As Boolean
    Editable = mProps.Editable
End Property

'Para poder saber si el elemento se imprime.  --->Agregado.  29/01/01  Angel P.
Public Property Let Imprimir(ByVal value As Boolean)
    mProps.Imprimir = value
    mbooModificado = True
End Property

Public Property Get Imprimir() As Boolean
    Imprimir = mProps.Imprimir
End Property

Public Property Let Orden(ByVal value As Long)
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "Elemento.Orden", MSGERR_INVALIDO
    mProps.Orden = value
    mbooModificado = True
End Property

Public Property Get Orden() As Long
    Orden = mProps.Orden
End Property

Public Property Let BandActivo(ByVal value As Boolean)
    mProps.BandActivo = value
    mbooModificado = True
End Property

Public Property Get BandActivo() As Boolean
    BandActivo = mProps.BandActivo
End Property

Public Property Let Visible(ByVal value As Boolean)
    mProps.Visible = value
    mbooModificado = True
End Property

Public Property Get Visible() As Boolean
    Visible = mProps.Visible
End Property

Public Property Let fechagrabado(ByVal value As Date)
    If Len(value) > 10 Then Err.Raise ERR_INVALIDO, "Elemento.FechaGrabado", MSGERR_INVALIDO
    mProps.fechagrabado = value
    mbooModificado = True
End Property

Public Property Get fechagrabado() As Date
    fechagrabado = mProps.fechagrabado
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

'*** METODOS
Public Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM vwElemento WHERE codElemento = '" & key & "'"
    Else
        sql = "SELECT * FROM vwElemento WHERE IdElemento = " & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        mProps.idElemento = !idElemento
        mProps.Codelemento = !Codelemento
        mProps.Nombre = !Nombre
        mProps.Descripcion = !Descripcion
        mProps.Formula = !Formula
        mProps.Meses = !Meses
        mProps.Tipo = !Tipo
        mProps.Editable = !Editable
        mProps.Imprimir = !Imprimir
        mProps.Orden = !Orden
        mProps.BandActivo = !BandActivo
        mProps.Visible = !Visible
        mProps.fechagrabado = !fechagrabado
        If Not IsNull(!bandValidarAsignar) Then mProps.ValidarAsignar = !bandValidarAsignar
        If Not IsNull(!bandmostrarenreporte) Then mProps.MostrarEnReporte = !bandmostrarenreporte
        If Not IsNull(!bandmostrarenprovision) Then mProps.MostrarEnprovision = !bandmostrarenprovision
        If Not IsNull(!BandAcumular) Then mProps.BandAcumular = !BandAcumular
        If Not IsNull(!AfectaEmp) Then mProps.BandAfectaEmp = !AfectaEmp
        If Not IsNull(!Debe) Then mProps.Debe = !Debe
        If Not IsNull(!Haber) Then mProps.Haber = !Haber
        If Not IsNull(!IdTipoRol) Then mProps.IdTipoRol = !IdTipoRol
        If Not IsNull(!CodTipoRol) Then mProps.CodTipoRol = !CodTipoRol
        If Not IsNull(!IdTipoRol1) Then mProps.IdTipoRol1 = !IdTipoRol1
        If Not IsNull(!CodTipoRol1) Then mProps.CodTipoRol1 = !CodTipoRol1
        If Not IsNull(!IdTipoRol2) Then mProps.IdTipoRol2 = !IdTipoRol2
        If Not IsNull(!CodTipoRol2) Then mProps.CodTipoRol2 = !CodTipoRol2
        If Not IsNull(!IdTipoRol3) Then mProps.IdTipoRol3 = !IdTipoRol3
        If Not IsNull(!CodTipoRol3) Then mProps.CodTipoRol3 = !CodTipoRol3
        If Not IsNull(!BandGastoDist) Then mProps.BandGastoDist = !BandGastoDist
        If Not IsNull(!BandBanco) Then mProps.BandBanco = !BandBanco
        'If Not IsNull(!IdTiposRol) Then mProps.IdTiposRol = !IdTiposRol
        If Not IsNull(!IdTipoRol4) Then mProps.IdTipoRol4 = !IdTipoRol4
        If Not IsNull(!CodTipoRol4) Then mProps.CodTipoRol4 = !CodTipoRol4
        If Not IsNull(!IdTipoRol5) Then mProps.IdTipoRol5 = !IdTipoRol5
        If Not IsNull(!CodTipoRol5) Then mProps.CodTipoRol5 = !CodTipoRol5
        If Not IsNull(!IdTipoRol6) Then mProps.IdTipoRol6 = !IdTipoRol6
        If Not IsNull(!CodTipoRol6) Then mProps.CodTipoRol6 = !CodTipoRol6
        If Not IsNull(!IdTipoRol7) Then mProps.IdTipoRol7 = !IdTipoRol7
        If Not IsNull(!CodTipoRol7) Then mProps.CodTipoRol7 = !CodTipoRol7
        If Not IsNull(!IdTipoRol8) Then mProps.IdTipoRol8 = !IdTipoRol8
        If Not IsNull(!CodTipoRol8) Then mProps.CodTipoRol8 = !CodTipoRol8
        If Not IsNull(!IdTipoRol9) Then mProps.IdTipoRol9 = !IdTipoRol9
        If Not IsNull(!CodTipoRol9) Then mProps.CodTipoRol9 = !CodTipoRol9
        If Not IsNull(!IdTipoRol10) Then mProps.IdTipoRol10 = !IdTipoRol10
        If Not IsNull(!CodTipoRol10) Then mProps.CodTipoRol10 = !CodTipoRol10
        If Not IsNull(!BandGastoDistPorc) Then mProps.BandGastoDistPorc = !BandGastoDistPorc
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Grabar()
    Dim sql As String, rs As Recordset
    On Error GoTo Etiq
    If Len(Me.Formula) <= 0 Then
        Err.Raise ERR_NOFORMULA, "Elemento.Grabar", MSGERR_NOFORMULA
    End If
    If mbooEsNuevo Then
        sql = "SELECT * FROM Elemento WHERE 1=0"
        Set rs = Me.Empresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    Else
        sql = "SELECT * FROM Elemento WHERE idElemento = '" & Me.idElemento & "'"
        Set rs = Me.Empresa.OpenRecordsetParaEdit(sql)
        
    End If
    With rs
        'rs!IdElemento = Me.IdElemento
        !Codelemento = Me.Codelemento
        !Nombre = Me.Nombre
        !Descripcion = Me.Descripcion
        If Not !Formula = mProps.Formula Then
            mProps.objEmpresa.GrabaAccionModulo !Formula, mProps.Formula, "Formula", "MOD_ELE", "GN"
        End If
        !Formula = Me.Formula
        !Meses = Me.Meses
        !Tipo = Me.Tipo
        !Editable = Me.Editable
        !Imprimir = Me.Imprimir
        !Orden = Me.Orden
        !BandActivo = Me.BandActivo
        !Visible = Me.Visible
        !fechagrabado = Now
        
        !bandValidarAsignar = Me.ValidarAsignar 'AUC 23/01/07
        !bandmostrarenreporte = Me.MostrarEnReporte 'AUC 23/02/07
        !bandmostrarenprovision = Me.MostrarEnprovision  'AUC 23/02/07
        !BandAcumular = Me.BandAcumular
        !AfectaEmp = Me.BandAfectaEmp
        !Debe = Me.Debe
        !Haber = Me.Haber
        !IdTipoRol = Me.IdTipoRol
        !IdTipoRol1 = Me.IdTipoRol1
        !IdTipoRol2 = Me.IdTipoRol2
        !IdTipoRol3 = Me.IdTipoRol3
        !IdTipoRol4 = Me.IdTipoRol4
        !IdTipoRol5 = Me.IdTipoRol5
        !IdTipoRol6 = Me.IdTipoRol6
        !IdTipoRol7 = Me.IdTipoRol7
        !IdTipoRol8 = Me.IdTipoRol8
        !IdTipoRol9 = Me.IdTipoRol9
        !IdTipoRol10 = Me.IdTipoRol10
        !BandGastoDist = Me.BandGastoDist
        !BandBanco = Me.BandBanco
        !BandGastoDistPorc = Me.BandGastoDistPorc
        .Update
        .Close
    End With
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_ELE", "Elemento Nuevo: " & _
                                                         mProps.Codelemento, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_ELE", "Elemento Modificado: " & _
                                                         mProps.Codelemento, "DT"
    End If
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    Me.Empresa.RaiseRefreshEvent REFRESH_ELEMENTOS
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
        sql = "DELETE FROM Elemento WHERE idElemento = '" & Me.idElemento & "'"
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        Me.Empresa.Coneccion.Execute sql
    mbooEsNuevo = True
    mbooModificado = True
    Me.Empresa.RaiseRefreshEvent REFRESH_ELEMENTOS
End Sub

'auc 23/01/2007
Public Property Let ValidarAsignar(ByVal value As Boolean)
    mProps.ValidarAsignar = value
    mbooModificado = True
End Property

Public Property Get ValidarAsignar() As Boolean
    ValidarAsignar = mProps.ValidarAsignar
End Property

'auc 23/02/2007
Public Property Let MostrarEnReporte(ByVal value As Boolean)
    mProps.MostrarEnReporte = value
    mbooModificado = True
End Property
Public Property Get MostrarEnReporte() As Boolean
    MostrarEnReporte = mProps.MostrarEnReporte
End Property

Public Property Let MostrarEnprovision(ByVal value As Boolean)
    mProps.MostrarEnprovision = value
    mbooModificado = True
End Property
Public Property Get MostrarEnprovision() As Boolean
    MostrarEnprovision = mProps.MostrarEnprovision
End Property

Public Property Let BandAcumular(ByVal value As Boolean)
    mProps.BandAcumular = value
    mbooModificado = True
End Property
Public Property Get BandAcumular() As Boolean
    BandAcumular = mProps.BandAcumular
End Property
Public Property Get idElemento() As Long
    idElemento = mProps.idElemento
End Property
Public Property Let Codelemento(ByVal value As String)
    If Len(value) > 10 Then Err.Raise ERR_INVALIDO, "Elemento.CodElemento", MSGERR_INVALIDO
    mProps.Codelemento = value
    mbooModificado = True
End Property

Public Property Get Codelemento() As String
    Codelemento = mProps.Codelemento
End Property

Public Property Let IdCuentaDebe(ByVal value As Long)
    mProps.IdCuentaDebe = value
    mbooModificado = True
End Property

Public Property Get IdCuentaDebe() As Long
    IdCuentaDebe = mProps.IdCuentaDebe
End Property


Public Property Let CodCuentaDebe(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodCuentaDebe Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Elementos_CodCuentaDebe", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaDebe = rs!IdCuenta
            End If
            rs.Close
        Else
            IdCuentaDebe = 0
        End If
        
        mProps.CodCuentaDebe = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property

Public Property Get CodCuentaDebe() As String
    CodCuentaDebe = Trim$(mProps.CodCuentaDebe)
End Property
Public Property Let IdCuentaHaber(ByVal value As Long)
    mProps.IdCuentaHaber = value
    mbooModificado = True
End Property

Public Property Get IdCuentaHaber() As Long
    IdCuentaHaber = mProps.IdCuentaHaber
End Property

Public Property Let CodCuentaHaber(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodCuentaHaber Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Elementos_CodCuentaHaber", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaHaber = rs!IdCuenta
            End If
            rs.Close
        Else
            IdCuentaHaber = 0
        End If
        mProps.CodCuentaHaber = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property

Public Property Get CodCuentaHaber() As String
    CodCuentaHaber = Trim$(mProps.CodCuentaHaber)
End Property

Public Property Let BandAfectaEmp(ByVal value As Boolean)
    mProps.BandAfectaEmp = value
    mbooModificado = True
End Property
Public Property Get BandAfectaEmp() As Boolean
    BandAfectaEmp = mProps.BandAfectaEmp
End Property

Public Property Get ExisteOrden() As Boolean
Dim sql As String
Dim rs As Recordset
On Error GoTo CapturaError
    sql = "select orden from elemento where orden =" & Me.Orden
    Set rs = Me.Empresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        ExisteOrden = True
        Exit Property
    End If
    Set rs = Nothing
    Exit Property
CapturaError:
    Set rs = Nothing
    MsgBox Err.Description
    Exit Property
End Property

Public Property Let Debe(ByVal value As Boolean)
    mProps.Debe = value
    mbooModificado = True
End Property

Public Property Get Debe() As Boolean
    Debe = mProps.Debe
End Property

Public Property Let Haber(ByVal value As Boolean)
    mProps.Haber = value
    mbooModificado = True
End Property

Public Property Get Haber() As Boolean
    Haber = mProps.Haber
End Property

Public Property Get IdTipoRol1() As Long
    IdTipoRol1 = mProps.IdTipoRol1
End Property
Public Property Let IdTipoRol1(ByVal value As Long)
    mProps.IdTipoRol1 = value
    mbooModificado = True
End Property
Public Property Let CodTipoRol1(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodTipoRol1 Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdTipo FROM tiporol WHERE Codigo='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Elementos_CodTipoRol1", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdTipoRol1 = rs!IdTipo
            End If
            rs.Close
        Else
            IdTipoRol1 = 0
        End If
        mProps.CodTipoRol1 = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property

Public Property Get CodTipoRol1() As String
    CodTipoRol1 = mProps.CodTipoRol1
End Property

Public Property Get IdTipoRol2() As Long
    IdTipoRol2 = mProps.IdTipoRol2
End Property
Public Property Let IdTipoRol2(ByVal value As Long)
    mProps.IdTipoRol2 = value
    mbooModificado = True
End Property
Public Property Let CodTipoRol2(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodTipoRol2 Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdTipo FROM tiporol WHERE Codigo='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Elementos_CodTipoRol2", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdTipoRol2 = rs!IdTipo
            End If
            rs.Close
        Else
            IdTipoRol2 = 0
        End If
        mProps.CodTipoRol2 = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property
Public Property Get CodTipoRol2() As String
    CodTipoRol2 = mProps.CodTipoRol2
End Property

Public Property Get IdTipoRol3() As Long
    IdTipoRol3 = mProps.IdTipoRol3
End Property
Public Property Let IdTipoRol3(ByVal value As Long)
    mProps.IdTipoRol3 = value
    mbooModificado = True
End Property
Public Property Let CodTipoRol3(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodTipoRol3 Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdTipo FROM tiporol WHERE Codigo='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Elementos_CodTipoRol3", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdTipoRol3 = rs!IdTipo
            End If
            rs.Close
        Else
            IdTipoRol3 = 0
        End If
        mProps.CodTipoRol3 = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property

Public Property Get CodTipoRol3() As String
    CodTipoRol3 = mProps.CodTipoRol3
End Property

Public Property Get IdTipoRol() As Long
    IdTipoRol = mProps.IdTipoRol
End Property
Public Property Let IdTipoRol(ByVal value As Long)
    mProps.IdTipoRol = value
    mbooModificado = True
End Property
Public Property Let CodTipoRol(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodTipoRol Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdTipo FROM tiporol WHERE Codigo='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Elementos_CodTipoRol", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdTipoRol = rs!IdTipo
            End If
            rs.Close
        Else
            IdTipoRol = 0
        End If
        mProps.CodTipoRol = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property
Public Property Get CodTipoRol() As String
    CodTipoRol = mProps.CodTipoRol
End Property
Public Property Let BandGastoDist(ByVal value As Boolean)
    mProps.BandGastoDist = value
    mbooModificado = True
End Property

Public Property Get BandGastoDist() As Boolean
    BandGastoDist = mProps.BandGastoDist
End Property

Public Property Let BandBanco(ByVal value As Boolean)
    mProps.BandBanco = value
    mbooModificado = True
End Property

Public Property Get BandBanco() As Boolean
    BandBanco = mProps.BandBanco
End Property


Public Property Get IdTipoRol4() As Long
    IdTipoRol4 = mProps.IdTipoRol4
End Property
Public Property Let IdTipoRol4(ByVal value As Long)
    mProps.IdTipoRol4 = value
    mbooModificado = True
End Property
Public Property Let CodTipoRol4(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodTipoRol4 Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdTipo FROM tiporol WHERE Codigo='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Elementos_CodTipoRol4", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdTipoRol4 = rs!IdTipo
            End If
            rs.Close
        Else
            IdTipoRol4 = 0
        End If
        mProps.CodTipoRol4 = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property

Public Property Get CodTipoRol4() As String
    CodTipoRol4 = mProps.CodTipoRol4
End Property

Public Property Get IdTipoRol5() As Long
    IdTipoRol5 = mProps.IdTipoRol5
End Property
Public Property Let IdTipoRol5(ByVal value As Long)
    mProps.IdTipoRol5 = value
    mbooModificado = True
End Property
Public Property Let CodTipoRol5(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodTipoRol5 Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdTipo FROM tiporol WHERE Codigo='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Elementos_CodTipoRol5", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdTipoRol5 = rs!IdTipo
            End If
            rs.Close
        Else
            IdTipoRol5 = 0
        End If
        mProps.CodTipoRol5 = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property

Public Property Get CodTipoRol5() As String
    CodTipoRol5 = mProps.CodTipoRol5
End Property

Public Property Get IdTipoRol6() As Long
    IdTipoRol6 = mProps.IdTipoRol6
End Property
Public Property Let IdTipoRol6(ByVal value As Long)
    mProps.IdTipoRol6 = value
    mbooModificado = True
End Property
Public Property Let CodTipoRol6(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodTipoRol6 Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdTipo FROM tiporol WHERE Codigo='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Elementos_CodTipoRol6", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdTipoRol6 = rs!IdTipo
            End If
            rs.Close
        Else
            IdTipoRol6 = 0
        End If
        mProps.CodTipoRol6 = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property

Public Property Get CodTipoRol6() As String
    CodTipoRol6 = mProps.CodTipoRol6
End Property

Public Property Get IdTipoRol7() As Long
    IdTipoRol7 = mProps.IdTipoRol7
End Property
Public Property Let IdTipoRol7(ByVal value As Long)
    mProps.IdTipoRol7 = value
    mbooModificado = True
End Property
Public Property Let CodTipoRol7(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodTipoRol7 Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdTipo FROM tiporol WHERE Codigo='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Elementos_CodTipoRol1", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdTipoRol7 = rs!IdTipo
            End If
            rs.Close
        Else
            IdTipoRol7 = 0
        End If
        mProps.CodTipoRol7 = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property

Public Property Get CodTipoRol7() As String
    CodTipoRol7 = mProps.CodTipoRol7
End Property

Public Property Get IdTipoRol8() As Long
    IdTipoRol8 = mProps.IdTipoRol8
End Property
Public Property Let IdTipoRol8(ByVal value As Long)
    mProps.IdTipoRol8 = value
    mbooModificado = True
End Property
Public Property Let CodTipoRol8(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodTipoRol8 Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdTipo FROM tiporol WHERE Codigo='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Elementos_CodTipoRol1", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdTipoRol8 = rs!IdTipo
            End If
            rs.Close
        Else
            IdTipoRol8 = 0
        End If
        mProps.CodTipoRol8 = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property

Public Property Get CodTipoRol8() As String
    CodTipoRol8 = mProps.CodTipoRol8
End Property

Public Property Get IdTipoRol9() As Long
    IdTipoRol9 = mProps.IdTipoRol9
End Property
Public Property Let IdTipoRol9(ByVal value As Long)
    mProps.IdTipoRol9 = value
    mbooModificado = True
End Property
Public Property Let CodTipoRol9(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodTipoRol9 Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdTipo FROM tiporol WHERE Codigo='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Elementos_CodTipoRol1", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdTipoRol9 = rs!IdTipo
            End If
            rs.Close
        Else
            IdTipoRol9 = 0
        End If
        mProps.CodTipoRol9 = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property

Public Property Get CodTipoRol9() As String
    CodTipoRol9 = mProps.CodTipoRol9
End Property

Public Property Get IdTipoRol10() As Long
    IdTipoRol10 = mProps.IdTipoRol10
End Property
Public Property Let IdTipoRol10(ByVal value As Long)
    mProps.IdTipoRol10 = value
    mbooModificado = True
End Property
Public Property Let CodTipoRol10(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodTipoRol1 Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdTipo FROM tiporol WHERE Codigo='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Elementos_CodTipoRol1", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdTipoRol10 = rs!IdTipo
            End If
            rs.Close
        Else
            IdTipoRol10 = 0
        End If
        mProps.CodTipoRol10 = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property

Public Property Get CodTipoRol10() As String
    CodTipoRol10 = mProps.CodTipoRol10
End Property

Public Sub RecuperarQuick(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT idElemento,AfectaEmp,Debe,Haber,IdTipoRol,IdTipoRol1,IdTipoRol2,IdTipoRol3,IdTipoRol4,IdTipoRol5,IdTipoRol6,IdTipoRol7,IdTipoRol8,IdTipoRol9,IdTipoRol10 FROM Elemento WHERE codElemento = '" & key & "'"
    Else
        sql = "SELECT idElemento,AfectaEmp,Debe,Haber,IdTipoRol,IdTipoRol1,IdTipoRol2,IdTipoRol3,IdTipoRol4,IdTipoRol5,IdTipoRol6,IdTipoRol7,IdTipoRol8,IdTipoRol9,IdTipoRol10 FROM Elemento WHERE IdElemento = " & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        mProps.idElemento = !idElemento
        If Not IsNull(!AfectaEmp) Then mProps.BandAfectaEmp = !AfectaEmp
        If Not IsNull(!Debe) Then mProps.Debe = !Debe
        If Not IsNull(!Haber) Then mProps.Haber = !Haber
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
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Property Let BandGastoDistPorc(ByVal value As Boolean)
    mProps.BandGastoDistPorc = value
    mbooModificado = True
End Property

Public Property Get BandGastoDistPorc() As Boolean
    BandGastoDistPorc = mProps.BandGastoDistPorc
End Property
