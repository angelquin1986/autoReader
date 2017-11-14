VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProdMaqMano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdMaqMano As Long
    CodMaqMano As String
    Descripcion As String
    BandMaquinaria As Boolean
    BandManoObra As Boolean
    CostoxHora As Currency
    LimiteHora As Currency
    FechaGrabado As Date
    Observacion As String         'Agregado AUC 31/08/2005
    BandValida As Boolean
    IdUnidad As Long
    CodUnidad As String
    
    objEmpresa As Empresa       'Referencia al objeto padre
    
End Type

Private mProps As T_PROPS
Private mcolContactos As Collection
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Public Sub Grabar()
    Dim sql As String, fgrabado As Date, i As Long
    Dim rs As Recordset
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
       
    'Verifica si no repite el código
    sql = "SELECT CodMaqMano FROM PRMaqMano " & _
          "WHERE (CodMaqMano='" & mProps.CodMaqMano & "') AND " & _
                "(IdMaqMano<>" & mProps.IdMaqMano & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "PCMatEmp", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    On Error GoTo ErrTrap

#If DAOLIB Then
    gobjWS.BeginTrans                           '======= Comienza la transacción
#Else
    mProps.objEmpresa.Coneccion.BeginTrans      '======= Comienza la transacción
#End If
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM PRMaqMano WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
    
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM PRMaqMano WHERE IdMaqMano=" & mProps.IdMaqMano
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "PCMatEmp_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en la base
    With rs
        !CodMaqMano = mProps.CodMaqMano
        !Descripcion = mProps.Descripcion
        !BandMaquinaria = mProps.BandMaquinaria
        !BandManoObra = mProps.BandManoObra
        !CostoxHora = mProps.CostoxHora
        !LimiteHora = mProps.LimiteHora
        !Observacion = mProps.Observacion  'Agregado AUC 31/08/2005
        !BandValida = mProps.BandValida
        !IdUnidad = mProps.IdUnidad 'auc
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        fgrabado = !FechaGrabado
        mProps.IdMaqMano = !IdMaqMano
        .Close
    End With
    
    
#If DAOLIB Then
    gobjWS.CommitTrans                  '======== Completa la transacción
#Else
    mProps.objEmpresa.Coneccion.CommitTrans      '======== Completa la transacción
#End If
    mProps.FechaGrabado = fgrabado
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_PC", "Proveedor/Cliente Nuevo: " & _
                                                         mProps.CodMaqMano, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_PC", "Proveedor/Cliente Modificado: " & _
                                                         mProps.CodMaqMano, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PRMAQMANO
    Exit Sub
ErrTrap:
#If DAOLIB Then
    gobjWS.Rollback                   '======== Anula la transacción
#Else
    mProps.objEmpresa.Coneccion.RollbackTrans  '======== Anula la transacción
#End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, i As Long
    Dim rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM VWPRMaqMano WHERE CodMaqMano='" & key & "'"
    Else
        sql = "SELECT * FROM VWPRMaqMano WHERE IdMaqMano=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IdMaqMano = !IdMaqMano
        mProps.CodMaqMano = !CodMaqMano
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!BandMaquinaria) Then mProps.BandMaquinaria = !BandMaquinaria
        If Not IsNull(!BandManoObra) Then mProps.BandManoObra = !BandManoObra
        If Not IsNull(!CostoxHora) Then mProps.CostoxHora = !CostoxHora
        If Not IsNull(!LimiteHora) Then mProps.LimiteHora = !LimiteHora
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!Observacion) Then mProps.Observacion = !Observacion      'Agregado AUC 31/08/2005
        If Not IsNull(!IdUnidad) Then mProps.IdUnidad = !IdUnidad
        If Not IsNull(!CodUnidad) Then mProps.CodUnidad = !CodUnidad
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
       
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset

    'Verifica si existe Kardex
    sql = "SELECT IdMaqMano FROM PCKardex WHERE IdMaqMano=" & mProps.IdMaqMano
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "PCMatEmp.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing

    'Elimina los contactos
    sql = "DELETE FROM PCContacto WHERE IdMaqMano=" & mProps.IdMaqMano
    mProps.objEmpresa.Execute sql, True
        
    'Elimina el registro
    sql = "DELETE FROM PCMatEmp WHERE CodMaqMano='" & mProps.CodMaqMano & "'"
    mProps.objEmpresa.Execute sql, True

    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_PC", "Proveedor/Cliente Eliminado: " & _
                                                         mProps.CodMaqMano, "DT"
                                                         
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PRMAQMANO
End Sub

Public Function Clone() As ProdMaqMano
    Dim nuevo As ProdMaqMano, nc As PCContacto, i As Long
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New ProdMaqMano
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
    Set nc = Nothing
End Function



Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdMaqMano = 0
    mbooModificado = True
End Property



Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
    
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Sub SetModificado()
    mbooModificado = True
End Sub


Public Property Get IdMaqMano() As Long
    IdMaqMano = mProps.IdMaqMano
End Property


Public Property Let CodMaqMano(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCMatEmp_CodMaqMano", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    '18-02-04  para controlar que el código no contenga espacio en blanco
    mProps.CodMaqMano = Trim(value)
    mbooModificado = True
End Property

Public Property Get CodMaqMano() As String
    CodMaqMano = Trim$(mProps.CodMaqMano)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCMatEmp_Descripcion", MSGERR_INVALIDO
    End If
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let BandMaquinaria(ByVal value As Boolean)
    If value <> mProps.BandMaquinaria Then
        mProps.BandMaquinaria = value
        mbooModificado = True
    End If
End Property

Public Property Get BandMaquinaria() As Boolean
    BandMaquinaria = mProps.BandMaquinaria
End Property

Public Property Let BandManoObra(ByVal value As Boolean)
    If value <> mProps.BandManoObra Then
        mProps.BandManoObra = value
        mbooModificado = True
    End If
End Property

Public Property Get BandManoObra() As Boolean
    BandManoObra = mProps.BandManoObra
End Property


Public Property Let CostoxHora(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "PCMatEmp_CostoxHora", MSGERR_INVALIDO
    End If
    
    mProps.CostoxHora = value
    mbooModificado = True
End Property

Public Property Get CostoxHora() As Currency
    CostoxHora = mProps.CostoxHora
End Property

Public Property Let LimiteHora(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "PCMatEmp_LimiteHora", MSGERR_INVALIDO
    End If
    
    mProps.LimiteHora = value
    mbooModificado = True
End Property

Public Property Get LimiteHora() As Currency
    LimiteHora = mProps.LimiteHora
End Property

Public Property Let BandValida(ByVal value As Boolean)
    If value <> mProps.BandValida Then
        mProps.BandValida = value
        mbooModificado = True
    End If
End Property

Public Property Get BandValida() As Boolean
    BandValida = mProps.BandValida
End Property

Public Property Get FechaGrabado() As Date
    FechaGrabado = mProps.FechaGrabado
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

'Agregado AUC 31/08/2005
Public Property Let Observacion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCMatEmp_Observacion", MSGERR_INVALIDO
    End If
    
    mProps.Observacion = value
    mbooModificado = True
End Property
'Agregado AUC 31/08/2005
Public Property Get Observacion() As String
    Observacion = Trim$(mProps.Observacion)
End Property


Public Property Let CodUnidad(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodUnidad Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdUnidad FROM ivUnidad WHERE CodUnidad='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVUnidad_CodUnidad", MSGERR_NOHAYCODIGO
                Exit Property
            Else  'ESTOY AQUI
                IdUnidad = rs!IdUnidad
            End If
            rs.Close
        Else
            IdUnidad = 0
        End If
        mProps.CodUnidad = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property
Public Property Get CodUnidad() As String
    CodUnidad = Trim$(mProps.CodUnidad)
End Property

Public Property Get IdUnidad() As Long
    IdUnidad = mProps.IdUnidad
End Property

Public Property Let IdUnidad(ByVal value As Long)
    mProps.IdUnidad = value
    mbooModificado = True
End Property

