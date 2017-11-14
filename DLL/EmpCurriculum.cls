VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmpCurriculum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type T_PROPS
    IdCurriculum As Long
    IdEmpleado As String
    NombreCon As String
    FechaNaciCon As Date
    CICon As String
    ActividadCon As String
    TrabajoCon As String
    DireccionTraCon As String
    TelefonoTraCon As String
    
    Primaria As String
    anoPrimaria As String
    Secundaria As String
    anoSecundaria As String
    Superior As String
    anoSuperior As String
    Titulo As String
    OtrosTitulo As String
    Conocimientos As String
    Cursos As String
    Experiencia As String
    NomRef1 As String
    TelRef1 As String
    EmPRef1 As String
    NomRef2 As String
    TelRef2 As String
    EmPRef2 As String
    NomRef3 As String
    TelRef3 As String
    EmPRef3 As String
    objPersonal As Personal
    mobjEmpresa As Empresa
End Type
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean
Private mProps As T_PROPS
Friend Property Set Empresa(ByVal obj As Empresa)
    Set mProps.mobjEmpresa = obj
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.mobjEmpresa
End Property

Public Property Get IdCurriculum() As Long
    IdCurriculum = mProps.IdCurriculum
End Property
Public Property Let IdCurriculum(ByVal value As Long)
    mProps.IdCurriculum = value
End Property

Public Property Get IdEmpleado() As String
    IdEmpleado = Trim$(mProps.IdEmpleado)
End Property

Public Property Let IdEmpleado(ByVal value As String)
    mProps.IdEmpleado = value
    mbooModificado = True
End Property


Private Sub Class_Terminate()
    Set mProps.objPersonal = Nothing
End Sub

Public Property Get Primaria() As String 'Propiedad  solo de lectura
    Primaria = Trim$(mProps.Primaria)
End Property
Public Property Let Primaria(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Personal_primaria", MSGERR_INVALIDO
    End If
     mProps.Primaria = value
    mbooModificado = True
End Property
 
Public Property Get anoPrimaria() As String 'Propiedad  solo de lectura
    anoPrimaria = Trim$(mProps.anoPrimaria)
End Property
Public Property Let anoPrimaria(ByVal value As String)
    If Len(value) > 4 Then
        Err.Raise ERR_INVALIDO, "Personal_anoprimaria", MSGERR_INVALIDO
    End If
     mProps.anoPrimaria = value
    mbooModificado = True
End Property
Public Property Get Secundaria() As String 'Propiedad  solo de lectura
    Secundaria = Trim$(mProps.Secundaria)
End Property
Public Property Let Secundaria(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Personal_secundaria", MSGERR_INVALIDO
    End If
     mProps.Secundaria = value
    mbooModificado = True
End Property

Public Property Get anoSecundaria() As String 'Propiedad  solo de lectura
    anoSecundaria = Trim$(mProps.anoSecundaria)
End Property
Public Property Let anoSecundaria(ByVal value As String)
    If Len(value) > 4 Then
        Err.Raise ERR_INVALIDO, "Personal_anosecundaria", MSGERR_INVALIDO
    End If
     mProps.anoSecundaria = value
    mbooModificado = True
End Property
Public Property Get Superior() As String 'Propiedad  solo de lectura
    Superior = Trim$(mProps.Superior)
End Property
Public Property Let Superior(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Personal_superior", MSGERR_INVALIDO
    End If
     mProps.Superior = value
    mbooModificado = True
End Property

Public Property Get anoSuperior() As String 'Propiedad  solo de lectura
    anoSuperior = Trim$(mProps.anoSuperior)
End Property
Public Property Let anoSuperior(ByVal value As String)
    If Len(value) > 4 Then
        Err.Raise ERR_INVALIDO, "Personal_anosuperior", MSGERR_INVALIDO
    End If
     mProps.anoSuperior = value
    mbooModificado = True
End Property

Public Property Get Titulo() As String 'Propiedad  solo de lectura
    Titulo = Trim$(mProps.Titulo)
End Property
Public Property Let Titulo(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Personal_titulo", MSGERR_INVALIDO
    End If
     mProps.Titulo = value
    mbooModificado = True
End Property
Public Property Get OtrosTitulo() As String 'Propiedad  solo de lectura
    OtrosTitulo = Trim$(mProps.OtrosTitulo)
End Property
Public Property Let OtrosTitulo(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Personal_Otrostitulo", MSGERR_INVALIDO
    End If
     mProps.OtrosTitulo = value
    mbooModificado = True
End Property

Public Property Get Conocimientos() As String 'Propiedad  solo de lectura
    Conocimientos = Trim$(mProps.Conocimientos)
End Property
Public Property Let Conocimientos(ByVal value As String)
    If Len(value) > 255 Then
        Err.Raise ERR_INVALIDO, "Personal_Conocimientos", MSGERR_INVALIDO
    End If
     mProps.Conocimientos = value
    mbooModificado = True
End Property

Public Property Get Cursos() As String 'Propiedad  solo de lectura
    Cursos = Trim$(mProps.Cursos)
End Property
Public Property Let Cursos(ByVal value As String)
    If Len(value) > 255 Then
        Err.Raise ERR_INVALIDO, "Personal_Cursos", MSGERR_INVALIDO
    End If
     mProps.Cursos = value
    mbooModificado = True
End Property

Public Property Get Experiencia() As String 'Propiedad  solo de lectura
    Experiencia = Trim$(mProps.Experiencia)
End Property
Public Property Let Experiencia(ByVal value As String)
    If Len(value) > 255 Then
        Err.Raise ERR_INVALIDO, "Personal_Experiencia", MSGERR_INVALIDO
    End If
     mProps.Experiencia = value
    mbooModificado = True
End Property

Public Property Get NomRef1() As String 'Propiedad  solo de lectura
    NomRef1 = Trim$(mProps.NomRef1)
End Property
Public Property Let NomRef1(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Personal_NomRef1", MSGERR_INVALIDO
    End If
     mProps.NomRef1 = value
    mbooModificado = True
End Property

Public Property Get TelRef1() As String 'Propiedad  solo de lectura
    TelRef1 = Trim$(mProps.TelRef1)
End Property
Public Property Let TelRef1(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "Personal_TelRef1", MSGERR_INVALIDO
    End If
     mProps.TelRef1 = value
    mbooModificado = True
End Property

Public Property Get EmPRef1() As String 'Propiedad  solo de lectura
    EmPRef1 = Trim$(mProps.EmPRef1)
End Property
Public Property Let EmPRef1(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Personal_EmPRef1", MSGERR_INVALIDO
    End If
     mProps.EmPRef1 = value
    mbooModificado = True
End Property

'-------------
Public Property Get NomRef2() As String 'Propiedad  solo de lectura
    NomRef2 = Trim$(mProps.NomRef2)
End Property
Public Property Let NomRef2(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Personal_NomRef2", MSGERR_INVALIDO
    End If
     mProps.NomRef2 = value
    mbooModificado = True
End Property

Public Property Get TelRef2() As String 'Propiedad  solo de lectura
    TelRef2 = Trim$(mProps.TelRef2)
End Property
Public Property Let TelRef2(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "Personal_TelRef2", MSGERR_INVALIDO
    End If
     mProps.TelRef2 = value
    mbooModificado = True
End Property

Public Property Get EmPRef2() As String 'Propiedad  solo de lectura
    EmPRef2 = Trim$(mProps.EmPRef2)
End Property
Public Property Let EmPRef2(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Personal_EmPRef2", MSGERR_INVALIDO
    End If
     mProps.EmPRef2 = value
    mbooModificado = True
End Property

'------------
Public Property Get NomRef3() As String 'Propiedad  solo de lectura
    NomRef3 = Trim$(mProps.NomRef3)
End Property
Public Property Let NomRef3(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Personal_NomRef3", MSGERR_INVALIDO
    End If
     mProps.NomRef3 = value
    mbooModificado = True
End Property

Public Property Get TelRef3() As String 'Propiedad  solo de lectura
    TelRef3 = Trim$(mProps.TelRef3)
End Property
Public Property Let TelRef3(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "Personal_TelRef3", MSGERR_INVALIDO
    End If
     mProps.TelRef3 = value
    mbooModificado = True
End Property

Public Property Get EmPRef3() As String 'Propiedad  solo de lectura
    EmPRef3 = Trim$(mProps.EmPRef3)
End Property
Public Property Let EmPRef3(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Personal_EmPRef3", MSGERR_INVALIDO
    End If
     mProps.EmPRef3 = value
    mbooModificado = True
End Property


Public Sub Recuperar(ByVal Id As Long)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT * FROM EmpCurriculum WHERE idEmpleado = " & Id
    
    Set rs = mProps.mobjEmpresa.OpenRecordset(sql)
        
    With rs
        Do While Not .EOF
            mProps.IdEmpleado = !IdEmpleado
            '------------
            If Not IsNull(!NombreCon) Then mProps.NombreCon = !NombreCon
            
            mProps.FechaNaciCon = IIf(IsNull(!FechaNaciCon), Date, !FechaNaciCon)
            
            If Not IsNull(!CICon) Then mProps.CICon = !CICon
            If Not IsNull(!ActividadCon) Then mProps.ActividadCon = !ActividadCon
            If Not IsNull(!TrabajoCon) Then mProps.TrabajoCon = !TrabajoCon
            If Not IsNull(!DireccionTraCon) Then mProps.DireccionTraCon = !DireccionTraCon
            If Not IsNull(!TelefonoTraCon) Then mProps.TelefonoTraCon = !TelefonoTraCon
            '------------
            If Not IsNull(!Primaria) Then mProps.Primaria = !Primaria
            If Not IsNull(!anoPrimaria) Then mProps.anoPrimaria = !anoPrimaria
            If Not IsNull(!Secundaria) Then mProps.Secundaria = !Secundaria
            If Not IsNull(!anoSecundaria) Then mProps.anoSecundaria = !anoSecundaria
            If Not IsNull(!Superior) Then mProps.Superior = !Superior
            If Not IsNull(!anoSuperior) Then mProps.anoSuperior = !anoSuperior
            If Not IsNull(!Titulo) Then mProps.Titulo = !Titulo
            If Not IsNull(!OtrosTitulo) Then mProps.OtrosTitulo = !OtrosTitulo
            If Not IsNull(!Conocimientos) Then mProps.Conocimientos = !Conocimientos
            If Not IsNull(!Cursos) Then mProps.Cursos = !Cursos
            If Not IsNull(!Experiencia) Then mProps.Experiencia = !Experiencia
            If Not IsNull(!NomRef1) Then mProps.NomRef1 = !NomRef1
            If Not IsNull(!TelRef1) Then mProps.TelRef1 = !TelRef1
            If Not IsNull(!EmPRef1) Then mProps.EmPRef1 = !EmPRef1
            If Not IsNull(!NomRef2) Then mProps.NomRef2 = !NomRef2
            If Not IsNull(!TelRef2) Then mProps.TelRef2 = !TelRef2
            If Not IsNull(!EmPRef2) Then mProps.EmPRef2 = !EmPRef2
            If Not IsNull(!NomRef3) Then mProps.NomRef3 = !NomRef3
            If Not IsNull(!TelRef3) Then mProps.TelRef3 = !TelRef3
            If Not IsNull(!EmPRef3) Then mProps.EmPRef3 = !EmPRef3
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
    
            sql = "SELECT * FROM EmpCurriculum WHERE idempleado = " & IdEmpleado
            Set rs = mProps.mobjEmpresa.OpenRecordsetParaEdit(sql)
                    
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            sql = "SELECT * FROM EmpCurriculum WHERE 1=0"
            Set rs = mProps.mobjEmpresa.OpenRecordsetParaEdit(sql)
            rs.AddNew
        Else
        End If
    
    With rs
            !IdEmpleado = IdEmpleado
            !NombreCon = mProps.NombreCon
            !FechaNaciCon = mProps.FechaNaciCon
            !CICon = mProps.CICon
            !ActividadCon = mProps.ActividadCon
            !TrabajoCon = mProps.TrabajoCon
            !DireccionTraCon = mProps.DireccionTraCon
            !TelefonoTraCon = mProps.TelefonoTraCon
            !Primaria = mProps.Primaria
            !anoPrimaria = mProps.anoPrimaria
            !Secundaria = mProps.Secundaria
            !anoSecundaria = mProps.anoSecundaria
             !Superior = mProps.Superior
             !anoSuperior = mProps.anoSuperior
             !Titulo = mProps.Titulo
             !OtrosTitulo = mProps.OtrosTitulo
             !Conocimientos = mProps.Conocimientos
             !Cursos = mProps.Cursos
             !Experiencia = mProps.Experiencia
             !NomRef1 = mProps.NomRef1
             !TelRef1 = mProps.TelRef1
             !EmPRef1 = mProps.EmPRef1
             !NomRef2 = mProps.NomRef2
             !TelRef2 = mProps.TelRef2
             !EmPRef2 = mProps.EmPRef2
             !NomRef3 = mProps.NomRef3
             !TelRef3 = mProps.TelRef3
             !EmPRef3 = mProps.EmPRef3
        .Update
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    Me.Empresa.RaiseRefreshEvent REFRESH_PROVCLI
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub


Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    
        sql = "DELETE FROM empCurriculum WHERE idEmpleado = " & Me.IdEmpleado
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        Me.Empresa.Coneccion.Execute sql
    
    mbooEsNuevo = True
    mbooModificado = True
    
End Sub

Public Property Get NombreCon() As String 'Propiedad  solo de lectura
    NombreCon = Trim$(mProps.NombreCon)
End Property
Public Property Let NombreCon(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Personal_NombreCon", MSGERR_INVALIDO
    End If
     mProps.NombreCon = value
    mbooModificado = True
End Property
Public Property Get FechaNaciCon() As Date
    FechaNaciCon = Trim$(mProps.FechaNaciCon)
End Property
Public Property Let FechaNaciCon(ByVal value As Date)
     mProps.FechaNaciCon = value
    mbooModificado = True
End Property
Public Property Get CICon() As String 'Propiedad  solo de lectura
    CICon = Trim$(mProps.CICon)
End Property
Public Property Let CICon(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "Personal_CICon", MSGERR_INVALIDO
    End If
     mProps.CICon = value
    mbooModificado = True
End Property
Public Property Get ActividadCon() As String 'Propiedad  solo de lectura
    ActividadCon = Trim$(mProps.ActividadCon)
End Property
Public Property Let ActividadCon(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Personal_ActividadCon", MSGERR_INVALIDO
    End If
     mProps.ActividadCon = value
    mbooModificado = True
End Property

Public Property Get TrabajoCon() As String 'Propiedad  solo de lectura
    TrabajoCon = Trim$(mProps.TrabajoCon)
End Property
Public Property Let TrabajoCon(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Personal_TrabajoCon", MSGERR_INVALIDO
    End If
     mProps.TrabajoCon = value
    mbooModificado = True
End Property
Public Property Get DireccionTraCon() As String 'Propiedad  solo de lectura
    DireccionTraCon = Trim$(mProps.DireccionTraCon)
End Property
Public Property Let DireccionTraCon(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Personal_DireccionTraCon", MSGERR_INVALIDO
    End If
     mProps.DireccionTraCon = value
    mbooModificado = True
End Property

Public Property Get TelefonoTraCon() As String 'Propiedad  solo de lectura
    TelefonoTraCon = Trim$(mProps.TelefonoTraCon)
End Property
Public Property Let TelefonoTraCon(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Personal_TelefonoTraCon", MSGERR_INVALIDO
    End If
     mProps.TelefonoTraCon = value
    mbooModificado = True
End Property
