VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImpuestoRenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    Id As Long
    Desde As Currency
    Hasta As Currency
    Fraccion As Currency
    Porcentaje As Single
    FechaAplica As Date
    objEmpresa As Empresa
End Type

Dim mProps As T_PROPS

Dim mbooModificado As Boolean
Dim mbooEsNuevo As Boolean

Public Property Get Id() As Long
    Id = mProps.Id
End Property

Public Property Get Porcentaje() As Single
    Porcentaje = mProps.Porcentaje
End Property

Public Property Let Porcentaje(ByVal value As Single)
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "ImpRentaEmp.Porcentaje", MSGERR_INVALIDO
    mProps.Porcentaje = value
    mbooModificado = True
End Property

Public Property Get FechaAplica() As Date
    FechaAplica = mProps.FechaAplica
End Property

Public Property Let FechaAplica(ByVal value As Date)
    If Not IsDate(value) Then Err.Raise ERR_INVALIDO, "ImpRentaEmp.FechaAplica", MSGERR_INVALIDO
    mProps.FechaAplica = value
    mbooModificado = True
End Property


Public Property Get Desde() As Currency
    Desde = mProps.Desde
End Property

Public Property Let Desde(ByVal value As Currency)
    mProps.Desde = value
    mbooModificado = True
End Property

Public Property Get Hasta() As Currency
    Hasta = mProps.Hasta
End Property

Public Property Let Hasta(ByVal value As Currency)
    mProps.Hasta = value
    mbooModificado = True
End Property

Public Property Get Fraccion() As Currency
    Fraccion = mProps.Fraccion
End Property

Public Property Let Fraccion(ByVal value As Currency)
    mProps.Fraccion = value
    mbooModificado = True
End Property

Public Property Get Modificado() As Boolean
    Modificado = mbooModificado
End Property

Public Property Set objEmpresa(ByVal value As Empresa)
    Set mProps.objEmpresa = value
End Property

Public Property Get objEmpresa() As Empresa
    Set objEmpresa = mProps.objEmpresa
End Property

Friend Sub Recuperar(Key As Long)
    Dim sql As String
    Dim rs As Recordset
    sql = "Select * from ImpuestoRenta Where IdIR = " & Key
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'Recupera los datos de Usuario
    With rs
        If Not rs.EOF Then
            mProps.Id = !idir
            mProps.Porcentaje = !Porcentaje
            If Not IsNull(!Desde) Then mProps.Desde = !Desde
            If Not IsNull(!Hasta) Then mProps.Hasta = !Hasta
            If Not IsNull(!Fraccion) Then mProps.Fraccion = !Fraccion
            If Not IsNull(!FechaAplica) Then mProps.FechaAplica = !FechaAplica
            mbooEsNuevo = False
            mbooModificado = True
        Else
            mProps.Porcentaje = 0
        End If
    End With
    Set rs = Nothing
End Sub



'Elimina en base a código de Porcentaje -> para casos en los que no se repite código, ejemplo tipo PRIVA    ALEX: FEB/2003
'sobre todo por el hecho de que ya se programó consultas que no devuelven id, y se complicaría más el cambio de ello en todo el código
Public Sub Eliminar()
    Dim sql As String, cond As String, TipoRefresh As String
    On Error GoTo ErrTrap
    sql = "DELETE FROM ImpuestoRenta WHERE Idir = " & mProps.Id
    mProps.objEmpresa.Execute sql
    mbooEsNuevo = True
    mbooModificado = True
    'Genera evento Refresh del objeto coaMain
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IR
    Exit Sub
ErrTrap:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Private Sub Class_Initialize()
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
'    mbooEsNuevo = True
'    mbooModificado = False
End Sub

Public Sub Grabar()
    Dim sql As String, rs As Recordset
    If Not mbooModificado Then Exit Sub
    'verificar si codigo está repetido
    'si es nuevo, agregar, si no modificar
    If Not mbooEsNuevo Then
        sql = "SELECT * FROM ImpuestoRenta WHERE idir = " & mProps.Id
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
    Else
        sql = "SELECT * FROM ImpuestoRenta WHERE idir = 0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    End If
    
    With rs
        !Desde = mProps.Desde
        !Hasta = mProps.Hasta
        !Fraccion = mProps.Fraccion
        !Porcentaje = mProps.Porcentaje
        !FechaAplica = mProps.FechaAplica
        .Update
        .Close
    End With
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_IR", "IR Nuevo: " & _
                                                         mProps.FechaAplica, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_IR", "IR Modificado: " & _
                                                         mProps.FechaAplica, "DT"
    End If
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IR
    Exit Sub
ErrTrap:

    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub



