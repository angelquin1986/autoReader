VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVEspProdISO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***jeaa  30/12/2004
Private Type T_PROPS
    Id As Long
    IdTamanio  As Long
    IdTrabajo  As Long
    Iddisenio  As Long
    CodTamanio  As String
    CodTrabajo  As String
    codDisenio  As String
    ValorPeso As Currency
    Precio As Currency
    Porcentaje As Currency
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
        
        
        sql = "select id "
        sql = sql & " from IVEspProdISO ive inner join ivinventario ivtam"
        sql = sql & " on ive.IdTamanio=ivtam.idinventario"
        sql = sql & " inner join ivinventario ivtra"
        sql = sql & " on ive.Idtrabajo=ivtra.idinventario"
        sql = sql & " inner join ivinventario ivdis"
        sql = sql & " on ive.Iddisenio=ivdis.idinventario"
        sql = sql & " WHERE (ivtam.codinventario='" & mProps.CodTamanio & "') AND "
        sql = sql & " (ivtra.codinventario='" & mProps.CodTrabajo & "') AND "
        sql = sql & " (ivdis.codinventario='" & mProps.codDisenio & "') AND "
        sql = sql & " (Id<>" & mProps.Id & ")"

    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "Especificacion Producto", MSGERR_REPITECODIGO
    End If
    rs.Close
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM IVEspProdISO WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM IVEspProdISO WHERE Id=" & mProps.Id
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)

        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "IVEspProdISO_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro
    With rs
        !IdTamanio = mProps.IdTamanio
        !IdTrabajo = mProps.IdTrabajo
        !Iddisenio = mProps.Iddisenio
        !ValorPeso = mProps.ValorPeso
        !Precio = mProps.Precio
        !Porcentaje = mProps.Porcentaje
        !FechaGrabado = Now
        .Update
       .Move 0             'Para actualizar
        mProps.FechaGrabado = !FechaGrabado
        mProps.Id = !Id
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVESPPRODISO
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String
    Dim rs As Recordset
    Dim v As Variant

    v = Split(key, ";")
        If VarType(key) = vbString Then
                sql = "SELECT id, ive.FechaGrabado, ValorPeso, Precio, Porcentaje, "
                sql = sql & " idtamanio, ivtam.codinventario as codtamanio ,"
                sql = sql & " idtrabajo, ivtra.codinventario as codtrabajo, "
                sql = sql & " iddisenio, ivdis.CodInventario as coddisenio "
                sql = sql & " from IVEspProdISO ive inner join ivinventario ivtam"
                sql = sql & " on ive.IdTamanio=ivtam.idinventario"
                sql = sql & " inner join ivinventario ivtra"
                sql = sql & " on ive.Idtrabajo=ivtra.idinventario"
                sql = sql & " inner join ivinventario ivdis"
                sql = sql & " on ive.Iddisenio=ivdis.idinventario"
                sql = sql & " WHERE ivtam.codinventario ='" & v(0) & "' "
                sql = sql & " and ivtra.codinventario='" & v(1) & "' "
                sql = sql & " and ivDIS.codinventario='" & v(2) & "'"
        
        End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'Recupera los datos de Usuario
    With rs
        mProps.Id = !Id
        mProps.IdTamanio = !IdTamanio
        mProps.CodTamanio = !CodTamanio
        mProps.IdTrabajo = !IdTrabajo
        mProps.CodTrabajo = !CodTrabajo
        mProps.Iddisenio = !Iddisenio
        'mProps.ValorPeso = !ValorPeso
        If Not IsNull(!ValorPeso) Then mProps.ValorPeso = !ValorPeso
        If Not IsNull(!Precio) Then mProps.Precio = !Precio
        If Not IsNull(!Porcentaje) Then mProps.Porcentaje = !Porcentaje
        mProps.codDisenio = !codDisenio
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
    sql = "DELETE FROM IVEspProdISO WHERE Id =" & mProps.Id
    mProps.objEmpresa.Execute sql, True

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVESPPRODISO
End Sub

Public Function Clone() As IVEspProdISO
    Dim nuevo As IVEspProdISO
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New IVEspProdISO
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Get Id() As Long
    Id = mProps.Id
End Property


Public Property Let IdTamanio(ByVal value As Long)
    mProps.IdTamanio = value
    mbooModificado = True
End Property

Public Property Get IdTamanio() As Long
    IdTamanio = mProps.IdTamanio
End Property

Public Property Let IdTrabajo(ByVal value As Long)
    mProps.IdTrabajo = value
    mbooModificado = True
End Property

Public Property Get IdTrabajo() As Long
    IdTrabajo = mProps.IdTrabajo
End Property

Public Property Let Iddisenio(ByVal value As Long)
    mProps.Iddisenio = value
    mbooModificado = True
End Property

Public Property Get Iddisenio() As Long
    Iddisenio = mProps.Iddisenio
End Property


Public Property Let CodTamanio(ByVal value As String)
    Dim sql As String, rs As Recordset
    Dim NumPCGrupo As Integer
'    NumPCGrupo = CInt(mProps.objEmpresa.GNOpcion.ObtenerValor("DescuentoNumPCGrupo")) + 1
    
    'Cuando cambia el valor
    If value <> mProps.CodTamanio Then
        sql = "SELECT IdInventario FROM IvInventario WHERE CodInventario='" & value & "'"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "CodTamanio", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            IdTamanio = rs.Fields("IdInventario")
        End If
        rs.Close
    End If
    
    Set rs = Nothing


    mProps.CodTamanio = value
    mbooModificado = True
End Property

Public Property Get CodTamanio() As String
    CodTamanio = mProps.CodTamanio
End Property

Public Property Let CodTrabajo(ByVal value As String)
    Dim sql As String, rs As Recordset
    Dim NumIVGrupo As Integer
    NumIVGrupo = CInt(mProps.objEmpresa.GNOpcion.ObtenerValor("DescuentoNumIVGrupo")) + 1
    
    'Cuando cambia el valor
    If value <> mProps.CodTrabajo Then
        sql = "SELECT IdInventario FROM IvInventario WHERE CodInventario='" & value & "'"

        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "CodTrabajo", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            IdTrabajo = rs.Fields("IdInventario")
        End If
        rs.Close
    End If
    
    Set rs = Nothing


    mProps.CodTrabajo = value
    mbooModificado = True
End Property

Public Property Get CodTrabajo() As String
    CodTrabajo = mProps.CodTrabajo
End Property


Public Property Let codDisenio(ByVal value As String)
    Dim sql As String, rs As Recordset
    Dim NumIVGrupo As Integer
    NumIVGrupo = CInt(mProps.objEmpresa.GNOpcion.ObtenerValor("DescuentoNumIVGrupo")) + 1
    
    'Cuando cambia el valor
    If value <> mProps.codDisenio Then
        sql = "SELECT IdInventario FROM IvInventario WHERE CodInventario='" & value & "'"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "Coddisenio", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            Iddisenio = rs.Fields("IdInventario")
        End If
        rs.Close
    End If
    
    Set rs = Nothing


    mProps.codDisenio = value
    mbooModificado = True
End Property

Public Property Get codDisenio() As String
    codDisenio = mProps.codDisenio
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
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Property Let ValorPeso(ByVal value As Currency)
    mProps.ValorPeso = value
    mbooModificado = True
End Property

Public Property Get ValorPeso() As Currency
    ValorPeso = mProps.ValorPeso
End Property

Public Property Let Precio(ByVal value As Currency)
    mProps.Precio = value
    mbooModificado = True
End Property

Public Property Get Precio() As Currency
    Precio = mProps.Precio
End Property

Public Property Let Porcentaje(ByVal value As Currency)
    mProps.Porcentaje = value
    mbooModificado = True
End Property

Public Property Get Porcentaje() As Currency
    Porcentaje = mProps.Porcentaje
End Property
