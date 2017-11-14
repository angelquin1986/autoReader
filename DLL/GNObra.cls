VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GNObra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Type T_PROPS
    IdObra As Long
    CodObra As String
    Nombre As String
    BandValida As Boolean
    BandPublica As Boolean
    FechaGrabado As Date
    IdDestino As Long
    CodDestino As String
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean


Public Sub Grabar()
    Dim sql As String, rs As Recordset

    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
    sql = "SELECT CodObra FROM GNObra " & _
          "WHERE (CodObra='" & mProps.CodObra & "') AND " & _
                "(IdObra<>" & mProps.IdObra & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "GNObra", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM GNObra WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM GNObra WHERE IdObra=" & mProps.IdObra
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "GNObra_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en Pos
    With rs
        !CodObra = mProps.CodObra
        !Nombre = mProps.Nombre
        !BandPublica = mProps.BandPublica
        !BandValida = mProps.BandValida
        !IdDestino = mProps.IdDestino
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.FechaGrabado = !FechaGrabado
        mProps.IdObra = !IdObra
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_OBRA", "Obra Nuevo: " & _
                                                         mProps.CodObra, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_OBRA", "Obra Modificado: " & _
                                                         mProps.CodObra, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_OBRA
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT gno.*, coddestino FROM GNObra gno left join gnDestino gnd on gno.iddestino=gnd.iddestino WHERE CodObra='" & key & "'"
    Else
        sql = "SELECT gno.*, coddestino FROM GNObra gno left join gnDestino gnd on gno.iddestino=gnd.iddestino WHERE IdObra=" & key
    End If

    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IdObra = !IdObra
        mProps.CodObra = !CodObra
        If Not IsNull(!Nombre) Then mProps.Nombre = !Nombre
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!BandPublica) Then mProps.BandPublica = !BandPublica
        If Not IsNull(!IdDestino) Then mProps.IdDestino = !IdDestino
        If Not IsNull(!CodDestino) Then mProps.CodDestino = !CodDestino
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    
    'Verifica si ya existe una transaccion usando el mismo
    sql = "SELECT IdObra FROM GNComprobante  WHERE IdObra=" & mProps.IdObra
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "GNObra.Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close

    'Elimina el registro
    sql = "DELETE FROM GNObra WHERE IdObra=" & mProps.IdObra
    mProps.objEmpresa.Execute sql, True

    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_OBRA", "Obra Eliminado: " & _
                                                         mProps.CodObra, "DT"

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_OBRA
End Sub

Friend Function Clone() As GNObra
    Dim nuevo As GNObra
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New GNObra
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdObra = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property






Public Property Let CodObra(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNObra_CodObra", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodObra = value
    mbooModificado = True
End Property

Public Property Get CodObra() As String
    CodObra = Trim$(mProps.CodObra)
End Property

Public Property Let Nombre(ByVal value As String)
    If Len(value) > 200 Then
        Err.Raise ERR_INVALIDO, "GNObra_Nombre", MSGERR_INVALIDO
    End If
    
    mProps.Nombre = value
    mbooModificado = True
End Property

Public Property Get Nombre() As String
    Nombre = Trim$(mProps.Nombre)
End Property

Public Property Let BandValida(ByVal value As Boolean)
    mProps.BandValida = value
    mbooModificado = True
End Property

Public Property Get BandValida() As Boolean
    BandValida = mProps.BandValida
End Property

Public Property Let BandPublica(ByVal value As Boolean)
    mProps.BandPublica = value
    mbooModificado = True
End Property

Public Property Get BandPublica() As Boolean
    BandPublica = mProps.BandPublica
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
    With mProps
        .BandValida = True
    End With
    
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub


Public Property Let IdDestino(ByVal value As Long)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.IdDestino Then
        'Actualiza el Codigo también
        mProps.CodDestino = ""
        sql = "SELECT CodDestino FROM Gndestino WHERE Iddestino=" & value
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then mProps.CodDestino = rs!CodDestino
        rs.Close
        Set rs = Nothing
    End If
    
    mProps.IdDestino = value
    mbooModificado = True
End Property

Public Property Get IdDestino() As Long
    IdDestino = mProps.IdDestino
End Property

Public Property Let CodDestino(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodDestino Then
        If Len(value) > 0 Then
            'Actualiza IdDestino también
            sql = "SELECT IdDestino FROM GNDestino WHERE CodDestino='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNObra.CodDestino", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdDestino = rs!IdDestino
            End If
            rs.Close
        Else
            mProps.IdDestino = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodDestino = value
    mbooModificado = True
End Property

Public Property Get CodDestino() As String
    CodDestino = Trim$(mProps.CodDestino)
End Property

