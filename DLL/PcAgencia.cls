VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCAgencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROP
    objPC As PCProvCli     'Referencia al objeto padre
    objSol As PCSolCreProvCli
    objEmpresa  As Empresa
    IdAgencia As Long
    IdProvcli As Long
    
    CodAgencia As String
    Descripcion As String
    IdCiudad As Long
    Direccion As String
    Ciudad As String
    Telefono As String
    Contacto As String
    Orden As Integer
    IdProvincia As Long
    CodProvincia As String
    
    IdVendedor As Long
    CodVendedor As String
    
    IdGrupo1 As Long
    CodGrupo1 As String
    
    IdGrupo2 As Long
    CodGrupo2 As String
    
    IdGrupo3 As Long
    CodGrupo3 As String
    
    IdGrupo4 As Long
    CodGrupo4 As String
    
    
    
    
    BandValida As Boolean
End Type
Private mProps As T_PROP

Dim mbooEsNuevo As Boolean
Dim mbooModificado As Boolean

Public Sub Grabar()
    Dim rs As Recordset, sql As String
    If Not mbooModificado Then Exit Sub
    'If mbooEsNuevo Then
    If mProps.IdAgencia = 0 Then
        sql = "SELECT * FROM PCAgencia WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    Else
         sql = "SELECT * FROM PCAgencia WHERE IdAGENCIA=" & mProps.IdAgencia
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
    End If
    With rs
        !IdProvcli = mProps.IdProvcli
        !IdCiudad = mProps.IdCiudad
        !IdProvincia = mProps.IdProvincia
        !CodAgencia = mProps.CodAgencia
        !Descripcion = mProps.Descripcion
        !Direccion = mProps.Direccion
        !Telefono = mProps.Telefono
        !Contacto = mProps.Contacto
        !Orden = mProps.Orden
        !BandValida = mProps.BandValida
        !IdVendedor = mProps.IdVendedor
        !IdGrupo1 = mProps.IdGrupo1
        !IdGrupo2 = mProps.IdGrupo2
        !IdGrupo3 = mProps.IdGrupo3
        !IdGrupo4 = mProps.IdGrupo4
        .Update
#If DAOLIB Then
        .Bookmark = .LastModified
#End If
        rs.Move 0
        mProps.IdAgencia = !IdAgencia
        .Close
    End With
    mbooEsNuevo = False
    mbooModificado = False
    Set rs = Nothing
End Sub

Friend Sub Recuperar(Id As Variant)
    Dim sql As String, rs As Recordset
    sql = "SELECT PCAgencia.*,pcc.codcanton as ciudad,pcv.codprovincia, "
    sql = sql & " pcg1.codgrupo1, pcg2.codgrupo2, pcg3.codgrupo3, pcg4.codgrupo4, fcv.codvendedor "
    sql = sql & " FROM PCAgencia left "
    sql = sql & " Join PCCanton pcc on pcc.idcanton = pcagencia.idciudad "
    sql = sql & " left Join PCProvincia pcv on pcv.idProvincia = pcagencia.idProvincia"
    sql = sql & " left join pcgrupo1 pcg1 on pcg1.idgrupo1 = PCAgencia.idgrupo1 "
    sql = sql & " left join pcgrupo2 pcg2 on pcg2.idgrupo2 = PCAgencia.idgrupo2 "
    sql = sql & " left join pcgrupo3 pcg3 on pcg3.idgrupo3 = PCAgencia.idgrupo3 "
    sql = sql & " left join pcgrupo4 pcg4 on pcg4.idgrupo4 = PCAgencia.idgrupo4 "
    sql = sql & " left join fcvendedor fcv on fcv.idvendedor= pcagencia.idvendedor "
    sql = sql & " Where IdAgencia = " & Id
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    With rs
        mProps.IdAgencia = !IdAgencia
        mProps.IdProvcli = !IdProvcli
        mProps.CodAgencia = !CodAgencia
        mProps.Descripcion = !Descripcion
        mProps.IdCiudad = !IdCiudad
        If Not IsNull(!Ciudad) Then mProps.Ciudad = !Ciudad
        mProps.IdProvincia = !IdProvincia
        If Not IsNull(!CodProvincia) Then mProps.CodProvincia = !CodProvincia
        
        If Not IsNull(!Telefono) Then mProps.Telefono = !Telefono
        If Not IsNull(!Direccion) Then mProps.Direccion = !Direccion
        If Not IsNull(!Contacto) Then mProps.Contacto = !Contacto
        If Not IsNull(!Orden) Then mProps.Orden = !Orden
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        
        If Not IsNull(!IdVendedor) Then mProps.IdVendedor = !IdVendedor
        If Not IsNull(!CodVendedor) Then mProps.CodVendedor = !CodVendedor
        If Not IsNull(!IdGrupo1) Then mProps.IdGrupo1 = !IdGrupo1
        If Not IsNull(!CodGrupo1) Then mProps.CodGrupo1 = !CodGrupo1
        If Not IsNull(!IdGrupo2) Then mProps.IdGrupo2 = !IdGrupo2
        If Not IsNull(!CodGrupo2) Then mProps.CodGrupo2 = !CodGrupo2
        If Not IsNull(!IdGrupo3) Then mProps.IdGrupo3 = !IdGrupo3
        If Not IsNull(!CodGrupo3) Then mProps.CodGrupo3 = !CodGrupo3
        If Not IsNull(!IdGrupo4) Then mProps.IdGrupo4 = !IdGrupo4
        If Not IsNull(!CodGrupo4) Then mProps.CodGrupo4 = !CodGrupo4
        
        .Close
    End With
    Set rs = Nothing
End Sub

Friend Function Clone() As PCAgencia
    Dim nuevo As PCAgencia
    
    Set nuevo = New PCAgencia
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROP)          'Para el método Clone
    mProps = value
    mbooModificado = True
End Property

Friend Property Set PCProvCli(value As PCProvCli)     'Para el método AddPCContacto de Grupo
    Set mProps.objPC = value
End Property

Public Property Get PCProvCli() As PCProvCli
    Set PCProvCli = mProps.objPC
End Property

Friend Property Set SolPC(value As PCSolCreProvCli)      'Para el método AddPCContacto de Grupo
    Set mProps.objSol = value
End Property

Public Property Get SolPC() As PCSolCreProvCli
    Set SolPC = mProps.objSol
End Property

Friend Property Set Empresa(value As Empresa)
    Set mProps.objEmpresa = value
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Get IdAgencia() As Long
    IdAgencia = mProps.IdAgencia
End Property

Public Property Let Contacto(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCAgencia_Contacto", MSGERR_INVALIDO
    End If
    mProps.Contacto = value
    mbooModificado = True
End Property

Public Property Get Contacto() As String
    Contacto = Trim$(mProps.Contacto)
End Property

Public Property Let Direccion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCAgencia_Direccion", MSGERR_INVALIDO
    End If
    mProps.Direccion = value
    mbooModificado = True
End Property

Public Property Get Direccion() As String
    Direccion = Trim$(mProps.Direccion)
End Property

Public Property Let Telefono(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "PCAgencia_Telefono", MSGERR_INVALIDO
    End If
    mProps.Telefono = value
    mbooModificado = True
End Property

Public Property Get Telefono() As String
    Telefono = Trim$(mProps.Telefono)
End Property

Public Property Let Orden(ByVal value As Integer)
    mProps.Orden = value
    mbooModificado = True
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property

Private Sub Class_Initialize()
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objPC = Nothing
    Set mProps.objEmpresa = Nothing
    mbooEsNuevo = False
End Sub

Public Property Let Ciudad(ByVal value As String)
 Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.Ciudad Then
        If Len(value) > 0 Then
            'Actualiza IdGrupo2 también
            sql = "SELECT IdCanton FROM PCCanton  WHERE CodCanton='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCAgencia_Ciudad", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCiudad = rs!IdCanton
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdCiudad = 0
        End If
    End If
    mProps.Ciudad = value
    mbooModificado = True
End Property

Public Property Get Ciudad() As String
    Ciudad = Trim$(mProps.Ciudad)
End Property

Public Property Let IdCiudad(ByVal value As Integer)
    mProps.IdCiudad = value
    mbooModificado = True
End Property

Public Property Get IdCiudad() As Integer
    IdCiudad = mProps.IdCiudad
End Property

Public Property Get EsNuevo() As Boolean
    EsNuevo = mbooEsNuevo
End Property

Public Property Get Modificado() As Boolean
    Modificado = mbooModificado
End Property

Friend Property Let Modificado(ByVal value As Boolean)
    mbooModificado = value
End Property
Public Sub EliminaFisica()
Dim sql As String
sql = "Delete from pcagencia where idagencia =" & Me.IdAgencia
    mProps.objEmpresa.EjecutarSQL sql, 1
    mbooEsNuevo = False
End Sub
Public Property Let CodAgencia(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCagencia_CodAgencia", MSGERR_INVALIDO
    End If
    mProps.CodAgencia = value
    mbooModificado = True
End Property

Public Property Get CodAgencia() As String
    CodAgencia = Trim$(mProps.CodAgencia)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCAgencia_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property
Public Property Let BandValida(ByVal value As Boolean)
    mProps.BandValida = value
    mbooModificado = True
End Property

Public Property Get BandValida() As Boolean
    BandValida = mProps.BandValida
End Property
Public Property Let IdProvcli(ByVal value As Integer)
    mProps.IdProvcli = value
    mbooModificado = True
End Property

Public Property Get IdProvcli() As Integer
    IdProvcli = mProps.IdProvcli
End Property
Friend Sub RecuperarXPC(Id As Variant)
'    Dim sql As String, rs As Recordset
'    sql = "SELECT PCAgencia.*,pcc.codcanton as ciudad FROM PCAgencia Inner Join PCCanton pcc on pcc.idcanton = pcagencia.idciudad "
'    sql = sql & " Where IdProvcli = " & Id
'    sql = sql & " And PCAgencia.BandValida =1"
'    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
'    With rs
'        mProps.IdAgencia = !IdAgencia
'        mProps.IdProvCli = !IdProvCli
'        mProps.CodAgencia = !CodAgencia
'        mProps.Descripcion = !Descripcion
'        mProps.IdCiudad = !IdCiudad
'        If Not IsNull(!Ciudad) Then mProps.Ciudad = !Ciudad
'        If Not IsNull(!Telefono) Then mProps.Telefono = !Telefono
'        If Not IsNull(!Direccion) Then mProps.Direccion = !Direccion
'        If Not IsNull(!Contacto) Then mProps.Contacto = !Contacto
'        If Not IsNull(!Orden) Then mProps.Orden = !Orden
'        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
'        .Close
'    End With
'    Set rs = Nothing
End Sub

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
    If value <> mProps.Ciudad Then
        If Len(value) > 0 Then
            'Actualiza IdGrupo2 también
            sql = "SELECT IdProvincia FROM PCProvincia  WHERE CodProvincia='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCAgencia_Provincia", MSGERR_NOHAYCODIGO
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
    mProps.CodProvincia = value
    mbooModificado = True
End Property

Public Property Get CodProvincia() As String
    CodProvincia = Trim$(mProps.CodProvincia)
End Property


Public Property Let IdVendedor(ByVal value As Long)
    mProps.IdVendedor = value
    mbooModificado = True
End Property

Public Property Get IdVendedor() As Long
    IdVendedor = mProps.IdVendedor
End Property


Public Property Let CodVendedor(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodVendedor Then
        If Len(value) > 0 Then
            'Actualiza IdVendedor también
            sql = "SELECT IdVendedor FROM FCVendedor WHERE CodVendedor='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodVendedor", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdVendedor = rs!IdVendedor
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdVendedor = 0
        End If
    End If

    mProps.CodVendedor = value
    mbooModificado = True
End Property

Public Property Get CodVendedor() As String
    CodVendedor = Trim$(mProps.CodVendedor)
End Property


Public Property Let IdGrupo1(ByVal value As Long)
    mProps.IdGrupo1 = value
    mbooModificado = True
End Property

Public Property Get IdGrupo1() As Long
    IdGrupo1 = mProps.IdGrupo1
End Property


Public Property Let CodGrupo1(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodGrupo1 Then
        If Len(value) > 0 Then
            'Actualiza IdGrupo1 también
            sql = "SELECT IdGrupo1 FROM PCGrupo1 WHERE CodGrupo1='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodGrupo1", MSGERR_NOHAYCODIGO
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
    mbooModificado = True
End Property

Public Property Get CodGrupo1() As String
    CodGrupo1 = Trim$(mProps.CodGrupo1)
End Property



Public Property Let IdGrupo2(ByVal value As Long)
    mProps.IdGrupo2 = value
    mbooModificado = True
End Property

Public Property Get IdGrupo2() As Long
    IdGrupo2 = mProps.IdGrupo2
End Property


Public Property Let CodGrupo2(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodGrupo2 Then
        If Len(value) > 0 Then
            'Actualiza IdGrupo2 también
            sql = "SELECT IdGrupo2 FROM PCGrupo2 WHERE CodGrupo2='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
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
    mbooModificado = True
End Property

Public Property Get CodGrupo2() As String
    CodGrupo2 = Trim$(mProps.CodGrupo2)
End Property


Public Property Let IdGrupo3(ByVal value As Long)
    mProps.IdGrupo3 = value
    mbooModificado = True
End Property
Public Property Let IdGrupo4(ByVal value As Long)
    mProps.IdGrupo4 = value
    mbooModificado = True
End Property

Public Property Get IdGrupo3() As Long
    IdGrupo3 = mProps.IdGrupo3
End Property


Public Property Let CodGrupo3(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodGrupo3 Then
        If Len(value) > 0 Then
            'Actualiza IdGrupo3 también
            sql = "SELECT IdGrupo3 FROM PCGrupo3 WHERE CodGrupo3='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodGrupo3", MSGERR_NOHAYCODIGO
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
    mbooModificado = True
End Property
Public Property Get IdGrupo4() As Long
    IdGrupo4 = mProps.IdGrupo4
End Property

Public Property Let CodGrupo4(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodGrupo4 Then
        If Len(value) > 0 Then
            'Actualiza IdGrupo3 también
            sql = "SELECT IdGrupo4 FROM PCGrupo4 WHERE CodGrupo4='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodGrupo4", MSGERR_NOHAYCODIGO
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
    mbooModificado = True
End Property

Public Property Get CodGrupo3() As String
    CodGrupo3 = Trim$(mProps.CodGrupo3)
End Property

Public Property Get CodGrupo4() As String
    CodGrupo4 = Trim$(mProps.CodGrupo4)
End Property

