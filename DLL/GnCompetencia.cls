VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GnCompetencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Type T_PROPS
    IdCompetencia As Long
    CodCompetencia As String
    Nombre As String
    BandValida As Boolean
    fechagrabado As Date
    'GARANTIA
    Direccion As String
    Telefono As String
    Valor As Currency
    FechaCaducidad As Date
    
    IdProvcli As Long
    CodProvcli As String
    'gualaTV
    Idzona As Long
    Codzona As String
    
    
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
    sql = "SELECT CodCompetencia FROM GNCompetencia " & _
          "WHERE (CodCompetencia='" & mProps.CodCompetencia & "') AND " & _
                "(IdCompetencia<>" & mProps.IdCompetencia & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "GNCompetencia", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM GNCompetencia WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM GNCompetencia WHERE IdCompetencia=" & mProps.IdCompetencia
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "GNCompetencia_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en Pos
    With rs
        !CodCompetencia = mProps.CodCompetencia
        !Nombre = mProps.Nombre
        !BandValida = mProps.BandValida
        
        !Direccion = mProps.Direccion
        !Telefono = mProps.Telefono
        !Valor = mProps.Valor
        !FechaCaducidad = mProps.FechaCaducidad
        !IdProvcli = mProps.IdProvcli
        
        !Idzona = mProps.Idzona
        
        
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IdCompetencia = !IdCompetencia
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_COM", "Competencia Nuevo: " & _
                                                         mProps.CodCompetencia, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_COM", "Competencia Modificado: " & _
                                                         mProps.CodCompetencia, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_GNCOMPETENCIA
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT gno.*, codCompetencia, codzona FROM GNCompetencia gno  LEFT JOIN GNZONA GNZ ON GNO.IDZONA= gnz.idzona WHERE CodCompetencia='" & key & "'"
    Else
        sql = "SELECT gno.*, codCompetencia FROM GNCompetencia gno  WHERE IdCompetencia=" & key
    End If

    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IdCompetencia = !IdCompetencia
        mProps.CodCompetencia = !CodCompetencia
        If Not IsNull(!Nombre) Then mProps.Nombre = !Nombre
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        
        If Not IsNull(!Direccion) Then mProps.Direccion = !Direccion
        If Not IsNull(!Telefono) Then mProps.Telefono = !Telefono
        If Not IsNull(!Valor) Then mProps.Valor = !Valor
        If Not IsNull(!FechaCaducidad) Then mProps.FechaCaducidad = !FechaCaducidad
        If Not IsNull(!IdProvcli) Then IdProvcli = !IdProvcli
        
        If Not IsNull(!Idzona) Then Idzona = !Idzona
        If Not IsNull(!Codzona) Then Codzona = !Codzona
        
        
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    
    'Verifica si ya existe una transaccion usando el mismo
    sql = "SELECT IdCompetencia FROM GNproyecto WHERE IdCompetencia=" & mProps.IdCompetencia
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "GNCompetencia.Eliminar", MSGERR_NOELIMINABLE
    End If
'    rs.Close


    sql = "SELECT IdsubZona FROM GNCentrocosto WHERE IdsubZona=" & mProps.IdCompetencia
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "GNCompetencia.Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close


    'Elimina el registro
    sql = "DELETE FROM GNCompetencia WHERE IdCompetencia=" & mProps.IdCompetencia
    mProps.objEmpresa.Execute sql, True

    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_COM", "Competencia Eliminado: " & _
                                                         mProps.CodCompetencia, "DT"

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_GNCOMPETENCIA
End Sub

Friend Function Clone() As GnCompetencia
    Dim nuevo As GnCompetencia
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New GnCompetencia
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdCompetencia = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property






Public Property Let CodCompetencia(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNCompetencia_CodCompetencia", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodCompetencia = value
    mbooModificado = True
End Property

Public Property Get CodCompetencia() As String
    CodCompetencia = Trim$(mProps.CodCompetencia)
End Property

Public Property Let Nombre(ByVal value As String)
    If Len(value) > 200 Then
        Err.Raise ERR_INVALIDO, "GNCompetencia_Nombre", MSGERR_INVALIDO
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
    With mProps
        .BandValida = True
        
    End With
    
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub


Public Property Let Direccion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "GNCompetencia_Direccion", MSGERR_INVALIDO
    End If
    
    mProps.Direccion = value
    mbooModificado = True
End Property

Public Property Get Direccion() As String
    Direccion = Trim$(mProps.Direccion)
End Property

Public Property Let Telefono(ByVal value As String)
    If Len(value) > 40 Then
        Err.Raise ERR_INVALIDO, "GNCompetencia_Telefono", MSGERR_INVALIDO
    End If
    
    mProps.Telefono = value
    mbooModificado = True
End Property

Public Property Get Telefono() As String
    Telefono = Trim$(mProps.Telefono)
End Property

Public Property Let Valor(ByVal value As Currency)

    mProps.Valor = value
    mbooModificado = True
End Property

Public Property Get Valor() As Currency
    Valor = Trim$(mProps.Valor)
End Property

Public Property Let FechaCaducidad(ByVal value As Date)
    mProps.FechaCaducidad = value
    mbooModificado = True
End Property

Public Property Get FechaCaducidad() As Date
    FechaCaducidad = Trim$(mProps.FechaCaducidad)
End Property

Public Property Let IdProvcli(ByVal value As Long)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.IdProvcli Then
        'Actualiza el Codigo también
        mProps.CodProvcli = ""
        sql = "SELECT CodProvCli FROM pcProvCli WHERE IdProvCli=" & value
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then mProps.CodProvcli = rs!CodProvcli
        rs.Close
        Set rs = Nothing
    End If
    
    mProps.IdProvcli = value
    mbooModificado = True
End Property

'*** MAKOTO 05/mar/01 Agregado
Public Property Get IdProvcli() As Long
    IdProvcli = mProps.IdProvcli
End Property

Public Property Let CodProvcli(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodProvcli Then
        If Len(value) > 0 Then
            'Actualiza IdProvCli también
            sql = "SELECT IdProvCli FROM pcProvCli WHERE CodProvCli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNContrato.CodProvCli", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdProvcli = rs!IdProvcli
            End If
            rs.Close
        Else
            mProps.IdProvcli = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodProvcli = value
    mbooModificado = True
End Property

Public Property Get CodProvcli() As String
    CodProvcli = Trim$(mProps.CodProvcli)
End Property

Public Property Let Idzona(ByVal value As Long)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.Idzona Then
        'Actualiza el Codigo también
        mProps.Codzona = ""
        sql = "SELECT Codzona FROM gnzona WHERE Idzona=" & value
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then mProps.Codzona = rs!Codzona
        rs.Close
        Set rs = Nothing
    End If
    
    mProps.Idzona = value
    mbooModificado = True
End Property

'*** MAKOTO 05/mar/01 Agregado
Public Property Get Idzona() As Long
    Idzona = mProps.Idzona
End Property

Public Property Let Codzona(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.Codzona Then
        If Len(value) > 0 Then
            'Actualiza Idzona también
            sql = "SELECT Idzona FROM gnzona WHERE Codzona='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNContrato.Codzona", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.Idzona = rs!Idzona
            End If
            rs.Close
        Else
            mProps.Idzona = 0
        End If
    End If

    Set rs = Nothing
    mProps.Codzona = value
    mbooModificado = True
End Property

Public Property Get Codzona() As String
    Codzona = Trim$(mProps.Codzona)
End Property


