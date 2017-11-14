VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVKardexSerie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    id As Long
    IdSerie As Long
    transid As Long
    IdBodega As Long
    Orden As Integer
    Cantidad As Currency
    CodBodega As String
    IdIvkardex As Long
    OrdenIvkardex As Integer
    objGNComprobante As GNComprobante
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Friend Sub Recuperar(rs As Recordset)
    Dim sql As String, i As Integer
    

'    If VarType(key) = vbString Then
'        sql = "SELECT * FROM IVSerie WHERE idSerie='" & key & "'"
'    Else
'        sql = "SELECT * FROM IVSerie WHERE idSerie=" & key
'    End If
'    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If rs.EOF Then Exit Sub
    'Recupera los datos de Usuario
    With rs
        mProps.IdSerie = !IdSerie
        mProps.transid = !transid
        mProps.Orden = !Orden
        mProps.Cantidad = !Cantidad
        mProps.IdBodega = !IdBodega
        mProps.CodBodega = !CodBodega
        mProps.IdIvkardex = !IdIvkardex
        mProps.OrdenIvkardex = !OrdenIvkardex
'        .Close
    End With
'    Set rs = Nothing
'    mbooEsNuevo = False
'    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    'Elimina el registro
    sql = "DELETE FROM IVKardexSerie WHERE Id=" & mProps.id
    mProps.objGNComprobante.Empresa.Execute sql, True
    mbooEsNuevo = True
    mbooModificado = True
End Sub

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdSerie = 0
    mbooModificado = True
End Property

Friend Property Set GNComprobante(value As GNComprobante)
    Set mProps.objGNComprobante = value
End Property

Public Property Get GNComprobante() As GNComprobante
    Set GNComprobante = mProps.objGNComprobante
End Property

Public Property Get IdSerie() As Long
    IdSerie = mProps.IdSerie
End Property

Public Property Let IdSerie(ByVal value As Long)
    mProps.IdSerie = value
    mbooModificado = True
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
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objGNComprobante = Nothing
End Sub

Public Property Get transid() As Long
    transid = mProps.transid
End Property
Public Property Let transid(ByVal value As Long)
    mProps.transid = value
    mbooModificado = True
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property
Public Property Let Orden(ByVal value As Integer)
    mProps.Orden = value
    mbooModificado = True
End Property

Public Sub GrabarIVKardexSerie(ByVal IdSerie As Long)
    Dim sql As String, i As Integer
    Dim rs As Recordset, rsDetalle As Recordset

    'Si no está modificado no hace nada
  '  If Not mbooModificado Then Exit Sub
    On Error GoTo ErrTrap
    AseguraBodega
    'Verifica si no repite el código
    'sql = "SELECT Serie FROM IVSerie " & _
          "WHERE idserie=" & mProps.IdSerie & _
          " And idinventario = " & mProps.IdInventario

'    Set rs = mProps.objIVKardex.GNComprobante.Empresa.OpenRecordset(sql)
 '   If Not rs.EOF Then      'Si ya existe el codigo
  '      Err.Raise ERR_REPITECODIGO, "IVInventarioSerie", MSGERR_REPITECODIGO
   ' End If
    'rs.Close
    'Set rs = Nothing
    
    'Si es nuevo
    If mbooEsNuevo Then
    'If objIVKardex.GNComprobante.EsNuevo Then
    'If idSerie = 0 Then
        sql = "SELECT * FROM IVKardexSerie WHERE 1=0"
        Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM IVKardexSerie WHERE Id=" & mProps.id
        Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    End If
    
    'Graba el registro
    With rs
        !IdSerie = mProps.IdSerie
        !transid = mProps.objGNComprobante.transid
        !IdBodega = mProps.IdBodega
        !IdIvkardex = mProps.IdIvkardex
        !Orden = mProps.Orden
        !Cantidad = mProps.Cantidad
        !OrdenIvkardex = mProps.OrdenIvkardex
        .Update
        rs.Move 0           'Para actualizar
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    Exit Sub
ErrTrap:
 MsgBox Err.Description
End Sub
Public Property Get IdBodega() As Long
    IdBodega = mProps.IdBodega
End Property
Public Property Let IdBodega(ByVal value As Long)
    mProps.IdBodega = value
    mbooModificado = True
End Property

Public Property Get Cantidad() As Currency
    Cantidad = mProps.Cantidad
End Property
Public Property Let Cantidad(ByVal value As Currency)
    mProps.Cantidad = value
    mbooModificado = True
End Property

Public Sub VerificaDatos()
    Dim iv As IVInventario
    Dim ivs As IVNumSerie
    Dim i As Long
    If mProps.Cantidad = 0 Then
        Err.Raise ERR_INVALIDO, "IVKNumSerie_VerificaDatos", _
            "La cantidad no puede ser 0. (" & mProps.Cantidad & ")"
    End If
    If mProps.IdBodega = 0 Then
        Err.Raise ERR_INVALIDO, "IVKNumSerie_VerificaDatos", _
            "Debe seleccionar una bodega. (" & mProps.IdBodega & ")"
    End If
   
    
End Sub


Private Sub AseguraBodega()
    Dim sql As String, rs As Recordset
    
    sql = "SELECT IdBodega FROM IVSerieExist WHERE IdSerie=" & _
                mProps.IdSerie & " AND IdBodega=" & mProps.IdBodega
    
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
    'Si no existe el registro correspondiente, lo inserta
    If rs.EOF Then
        sql = "INSERT INTO IVSerieExist (IdSerie,IdBodega,Exist) " & _
              "VALUES(" & mProps.IdSerie & "," & mProps.IdBodega & "," & 0 & ")"
        mProps.objGNComprobante.Empresa.Execute sql, True
    End If
    rs.Close
    Set rs = Nothing
End Sub


Public Property Let CodBodega(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodBodega Then
        If Len(value) > 0 Then
            'Actualiza IdBodega también
            sql = "SELECT IdBodega FROM IVBodega WHERE CodBodega='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVKardexSerie_CodBodega", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                Me.IdBodega = rs!IdBodega
            End If
            rs.Close
        Else
            Me.IdBodega = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodBodega = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodBodega() As String
    CodBodega = Trim$(mProps.CodBodega)
End Property

Public Property Get IdIvkardex() As Long
    IdIvkardex = mProps.IdIvkardex
End Property

Public Property Let IdIvkardex(ByVal value As Long)
    mProps.IdIvkardex = value
    mbooModificado = True
End Property

Public Property Get OrdenIvkardex() As Integer
    OrdenIvkardex = mProps.OrdenIvkardex
End Property
Public Property Let OrdenIvkardex(ByVal value As Integer)
    mProps.OrdenIvkardex = value
    mbooModificado = True
End Property
Friend Function Clone() As IVKardexSerie
    Dim nuevo As IVKardexSerie
    
    Set nuevo = New IVKardexSerie
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function
