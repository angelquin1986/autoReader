VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVDescuento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***jeaa  30/12/2004
Private Type T_PROPS
    
    CodDescto As String
    Descripcion As String
    BandValida As Boolean
    BandCaducidad As Boolean
    FechaDesde As Date
    FechaHasta As Date
    HoraDesde As Date
    HoraHasta As Date
    
    FechaGrabado As Date
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
        
        sql = "SELECT Valor FROM IVDescuento " & _
          "WHERE (CodPCGrupo='" & mProps.CodPCGrupo & "') AND " & _
          " (codIVGrupo='" & mProps.CodIVGrupo & "') AND " & _
                "(IdDescuento<>" & mProps.IdDescuento & ")"

    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "Descuento PCGrupo x IVGrupo", MSGERR_REPITECODIGO
    End If
    rs.Close
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM IVDescuento WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM IVDescuento WHERE IdDescuento=" & mProps.IdDescuento
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)

        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "DesctPCGxIVG_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro
    With rs
        'jeaa 05/10/2010
        !CodDescto = mProps.CodDescto
        !Descripcion = mProps.Descripcion
        !BandValida = mProps.BandValida
        !BandCaducidad = mProps.BandCaducidad
        !FechaDesde = mProps.FechaDesde
        !FechaHasta = mProps.FechaHasta
        !HoraDesde = mProps.HoraDesde
        !HoraHasta = mProps.HoraHasta
        
        
        !FechaGrabado = Now
        .Update
        .Move 0             'Para actualizar
        mProps.FechaGrabado = !FechaGrabado
        mProps.IdDescuento = !IdDescuento
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_DPCGXIVG
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String
    Dim rs As Recordset
    Dim v As Variant
    Dim NumIVGrupo As Integer, NumPCGrupo As Integer
    NumIVGrupo = CInt(mProps.objEmpresa.GNOpcion.ObtenerValor("DescuentoNumIVGrupo")) + 1
    NumPCGrupo = CInt(mProps.objEmpresa.GNOpcion.ObtenerValor("DescuentoNumPCGrupo")) + 1

    v = Split(key, ",")
        If VarType(key) = vbString Then
            sql = "SELECT *"
            sql = sql & " FROM IVDescuento "
            sql = sql & " WHERE CodDescto='" & key & "'"
        End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'Recupera los datos de Usuario
    With rs
        
        If Not IsNull(!CodDescto) Then mProps.CodDescto = !CodDescto
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        mProps.BandValida = !BandValida
        mProps.BandCaducidad = !BandCaducidad
        mProps.FechaDesde = !FechaDesde
        mProps.FechaHasta = !FechaHasta
        mProps.HoraDesde = !HoraDesde
        mProps.HoraHasta = !HoraHasta
        mProps.FechaGrabado = !FechaGrabado
        
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
  
    'Elimina el registro
    sql = "DELETE FROM IVDescuento WHERE Iddescuento=" & mProps.IdDescuento
    mProps.objEmpresa.Execute sql, True

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVDESCUENTO
End Sub

Public Function Clone() As IVDescuento
    Dim nuevo As IVDescuento
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New IVDescuento
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdDescuento = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Get IdDescuento() As Long
    IdDescuento = mProps.IdDescuento
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

Friend Property Let Modificado(ByVal value As Boolean)
    mbooModificado = value
End Property


Private Sub Class_Initialize()
    'Valores predeterminados
    mbooEsNuevo = True
    mbooModificado = False
    mProps.BandValida = 1
    
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Property Let CodDescto(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "IVDescuento_CodDescto", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodDescto = value
    mbooModificado = True
End Property

Public Property Get CodDescto() As String
    CodDescto = Trim$(mProps.CodDescto)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 254 Then
        Err.Raise ERR_INVALIDO, "IVDescuento_Descripcion", MSGERR_INVALIDO
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

Public Property Let BandCaducidad(ByVal value As Boolean)
    mProps.BandCaducidad = value
    mbooModificado = True
End Property

Public Property Get BandCaducidad() As Boolean
    BandCaducidad = mProps.BandCaducidad
End Property

Public Property Let FechaDesde(value As Date)
        mProps.FechaDesde = value
        mbooModificado = True
End Property

Public Property Get FechaDesde() As Date
    FechaDesde = mProps.FechaDesde
End Property

Public Property Let FechaHasta(value As Date)
        mProps.FechaHasta = value
        mbooModificado = True
End Property

Public Property Get FechaHasta() As Date
    FechaHasta = mProps.FechaHasta
End Property


Public Property Let HoraDesde(value As Date)
        mProps.HoraDesde = value
        mbooModificado = True
End Property

Public Property Get HoraDesde() As Date
    HoraDesde = mProps.HoraDesde
End Property

Public Property Let HoraHasta(value As Date)
        mProps.HoraHasta = value
        mbooModificado = True
End Property

Public Property Get HoraHasta() As Date
    HoraHasta = mProps.HoraHasta
End Property





