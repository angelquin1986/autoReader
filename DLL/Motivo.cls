VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Motivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'jeaa 25/04/2005
Option Explicit
Private Type T_PROPS
    IdMotivo As Long
    CodMotivo As String
    Descripcion As String
    Descripcion2 As String
    BandValida As Boolean
        
    IdCliente As Long
    CodCliente As String
    BandArribo As Boolean
    IdDestino As Long
    CodDestino As String

    
   fechagrabado As Date
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
    sql = "SELECT CodMotivo FROM Motivo " & _
          " WHERE (CodMotivo)='" & mProps.CodMotivo & "' AND " & _
                 "(IdMotivo <>" & mProps.IdMotivo & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "IVMotivo_Grabar", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM Motivo WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM Motivo WHERE IdMotivo =" & mProps.IdMotivo
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "IVMotivo_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en Pos
    With rs
        !CodMotivo = mProps.CodMotivo
        !Descripcion = mProps.Descripcion
        !Descripcion2 = mProps.Descripcion2
        !BandValida = mProps.BandValida
        !IdCliente = mProps.IdCliente
        !BandArribo = mProps.BandArribo
        !IdDestino = mProps.IdDestino
        
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IdMotivo = !IdMotivo
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_MOT", "Motivo Nuevo: " & _
                                                         mProps.CodMotivo, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_MOT", "Motivo Modificado: " & _
                                                         mProps.CodMotivo, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_MOTIVO
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM Motivo WHERE CodMotivo='" & key & "'"
    Else
        sql = "SELECT * FROM Motivo WHERE IdMotivo=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Motivo
    With rs
        mProps.IdMotivo = !IdMotivo
        mProps.CodMotivo = !CodMotivo
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!Descripcion2) Then mProps.Descripcion2 = !Descripcion2
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!IdCliente) Then IdCliente = !IdCliente
        If Not IsNull(!BandArribo) Then mProps.BandArribo = !BandArribo
        If Not IsNull(!IdDestino) Then IdDestino = !IdDestino
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String
    Dim rs As Recordset
    
    'Verifica si ya existe un item usando el mismo CodMotivo
    sql = "SELECT CodTrans FROM GnComprobante" & _
          " WHERE IdMotivo=" & mProps.IdMotivo
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "Motivo_Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close

    'Elimina el registro
    sql = "DELETE FROM Motivo WHERE IdMotivo=" & mProps.IdMotivo
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_MOT", "Motivo Eliminado: " & _
                                                         mProps.CodMotivo, "DT"
                                                         
    Set rs = Nothing
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_MOTIVO
End Sub

Friend Function Clone() As Motivo
    Dim nuevo As Motivo
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New Motivo
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdMotivo = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Let CodMotivo(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "Motivo_CodMotivo", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodMotivo = value
    mbooModificado = True
End Property

Public Property Get CodMotivo() As String
    CodMotivo = Trim$(mProps.CodMotivo)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "Motivo_Descripcion", MSGERR_INVALIDO
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
    
    'Valores iniciales
    mProps.BandValida = True
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Property Let IdCliente(ByVal value As Long)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.IdCliente Then
        'Actualiza el Codigo también
        mProps.CodCliente = ""
        sql = "SELECT CodProvCli FROM PCProvCli WHERE IdProvCli=" & value
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then mProps.CodCliente = rs!CodProvCli
        rs.Close
        Set rs = Nothing
    End If
    
    mProps.IdCliente = value
    mbooModificado = True
End Property

'*** MAKOTO 05/mar/01 Agregado
Public Property Get IdCliente() As Long
    IdCliente = mProps.IdCliente
End Property

Public Property Let CodCliente(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCliente Then
        If Len(value) > 0 Then
            'Actualiza IdCliente también
            sql = "SELECT IdProvCli FROM PCProvCli WHERE CodProvCli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNCentroCosto.CodCliente", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdCliente = rs!IdProvCli
            End If
            rs.Close
        Else
            mProps.IdCliente = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodCliente = value
    mbooModificado = True
End Property

Public Property Get CodCliente() As String
    CodCliente = Trim$(mProps.CodCliente)
End Property



    
Public Property Let BandArribo(ByVal value As Boolean)
    mProps.BandArribo = value
    mbooModificado = True
End Property

Public Property Get BandArribo() As Boolean
    BandArribo = mProps.BandArribo
End Property

Public Property Let Descripcion2(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "Motivo_Descripcion2", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion2 = value
    mbooModificado = True
End Property

Public Property Get Descripcion2() As String
    Descripcion2 = Trim$(mProps.Descripcion2)
End Property

Public Property Let IdDestino(ByVal value As Long)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.IdDestino Then
        'Actualiza el Codigo también
        mProps.CodDestino = ""
        sql = "SELECT CodDestino FROM GnDestino WHERE IdDestino=" & value
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then mProps.CodDestino = rs!CodDestino
        rs.Close
        Set rs = Nothing
    End If
    
    mProps.IdDestino = value
    mbooModificado = True
End Property

'*** MAKOTO 05/mar/01 Agregado
Public Property Get IdDestino() As Long
    IdDestino = mProps.IdDestino
End Property

Public Property Let CodDestino(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodDestino Then
        If Len(value) > 0 Then
            'Actualiza IdDestino también
            sql = "SELECT IdDestino FROM GnDestino WHERE CodDestino='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNDestino.CodDestino", MSGERR_NOHAYCODIGO
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


Public Property Get IdMotivo() As Long
    IdMotivo = mProps.IdMotivo
End Property

