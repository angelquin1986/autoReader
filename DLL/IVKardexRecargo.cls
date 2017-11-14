VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVKardexRecargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"GNComprobante"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Private Type T_PROP
    Id As Long
    IdRecargo As Long
    CodRecargo As String
    Porcentaje As Currency
    Valor As Currency
    BandModificable As Boolean
    BandOrigen As Integer
    BandProrrateado As Boolean
    AfectaIvaItem As Boolean        '*** MAKOTO 13/dic/00
    Orden As Integer
    BandSeleccionable  As Boolean
    auxSigno As Integer
    auxDescripcion As String
    objGNComprobante As GNComprobante           'Referencia al objeto padre
End Type

Private mProps As T_PROP

Public Sub VerificaDatos()
    'Verifica datos
    If mProps.IdRecargo = 0 And mProps.BandOrigen <> REC_SUBTOTAL Then
        Err.Raise ERR_INVALIDO, "IVKardexRecargo_VerificaDatos", _
            "Código de recargo no está asignado."
    End If
End Sub

Friend Sub Grabar()
    Dim sql As String, rs As Recordset
    
    VerificaDatos
    
    '------
    'Aquí podría poner If para que no grabe si no tiene valor (Pendiente)
    '------
    
    sql = "SELECT * FROM IVKardexRecargo WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !TransID = mProps.objGNComprobante.TransID
        !IdRecargo = mProps.IdRecargo
        !Porcentaje = mProps.Porcentaje
        !Valor = mProps.Valor
        !BandModificable = mProps.BandModificable
        !BandOrigen = mProps.BandOrigen
        !BandProrrateado = mProps.BandProrrateado
        !AfectaIvaItem = mProps.AfectaIvaItem           '*** MAKOTO 13/dic/00
        !Orden = mProps.Orden
        .Update
        
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.Id = !Id             'Guarda el Id(Autonumerico)
        .Close
    End With
    Set rs = Nothing
End Sub

Friend Sub Recuperar(rs As Recordset)
    Dim sql As String, rs1 As Recordset
    With rs
        If Not .EOF Then
            mProps.Id = !Id
            If Not IsNull(!IdRecargo) Then mProps.IdRecargo = !IdRecargo
            If Not IsNull(!CodRecargo) Then mProps.CodRecargo = !CodRecargo
            mProps.Porcentaje = !Porcentaje
            mProps.Valor = !Valor
            mProps.BandModificable = !BandModificable
            mProps.BandOrigen = !BandOrigen
            mProps.BandProrrateado = !BandProrrateado
            If Not IsNull(!AfectaIvaItem) Then mProps.AfectaIvaItem = !AfectaIvaItem    '*** MAKOTO 13/dic/00
            mProps.Orden = !Orden
        End If
    End With
End Sub

Friend Function Clone() As IVKardexRecargo
    Dim nuevo As IVKardexRecargo
    
    Set nuevo = New IVKardexRecargo
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROP)          'Para el método Clone
    mProps = value
    mProps.Id = 0
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Set GNComprobante(value As GNComprobante)
    Set mProps.objGNComprobante = value
End Property

Public Property Get GNComprobante() As GNComprobante
    Set GNComprobante = mProps.objGNComprobante
End Property



Public Property Get Id() As Long
    Id = mProps.Id
End Property


Public Sub SetCodRecargoAux(Signo As Integer)
    mProps.auxSigno = Signo
End Sub

Public Property Let IdRecargo(ByVal value As Long)
    mProps.IdRecargo = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get IdRecargo() As Long
    IdRecargo = mProps.IdRecargo
End Property


Public Property Let CodRecargo(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodRecargo Then
        'Actualiza IdRecargo también
        sql = "SELECT IdRecargo FROM IVRecargo WHERE CodRecargo='" & value & "'"
        Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "IVKardexRecargo_CodRecargo", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            IdRecargo = rs!IdRecargo
        End If
        rs.Close
    End If

    Set rs = Nothing
    mProps.CodRecargo = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodRecargo() As String
    CodRecargo = Trim$(mProps.CodRecargo)
End Property



Public Property Let Porcentaje(ByVal value As Currency)
    mProps.Porcentaje = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Porcentaje() As Currency
    Porcentaje = mProps.Porcentaje
End Property

Public Property Let Valor(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    
    mProps.Valor = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Valor() As Currency
    'Devuelve en Moneda Ex.
    Valor = mProps.Valor / mProps.objGNComprobante.Cotizacion("")
End Property

Public Property Let BandModificable(ByVal value As Boolean)
    mProps.BandModificable = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get BandModificable() As Boolean
    BandModificable = mProps.BandModificable
End Property

'***Diego 13/10/2003
Public Property Get BandSeleccionable() As Boolean
    Dim sql As String, rs1 As Recordset

    'carga  la propiedad configurada
    sql = "SELECT BandSeleccionable FROM GNTransRecargo WHERE IDRecargo=" & mProps.IdRecargo & _
          " AND CodTrans = '" & mProps.objGNComprobante.CodTrans & "'"
    Set rs1 = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
    If rs1.EOF Then
        Err.Raise ERR_NOHAYCODIGO, "IVKardexRecargo_BandSeleccionable", MSGERR_NOHAYCODIGO
        Exit Sub
    Else
        If IsNull(rs1!BandSeleccionable) Then
            mProps.BandSeleccionable = False
        Else
            mProps.BandSeleccionable = rs1!BandSeleccionable
        End If
        'BandSeleccionable = mProps.BandModificable
    End If
    rs1.Close
    Set rs1 = Nothing
    BandSeleccionable = mProps.BandSeleccionable
End Property

Public Property Let BandOrigen(ByVal value As Integer)
    mProps.BandOrigen = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get BandOrigen() As Integer
    BandOrigen = mProps.BandOrigen
End Property

Public Property Let BandProrrateado(ByVal value As Boolean)
    mProps.BandProrrateado = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get BandProrrateado() As Boolean
    BandProrrateado = mProps.BandProrrateado
End Property

'*** MAKOTO 13/dic/00 Agregado
Public Property Let AfectaIvaItem(ByVal value As Boolean)
    mProps.AfectaIvaItem = value
    mProps.objGNComprobante.SetModificado
End Property

'*** MAKOTO 13/dic/00 Agregado
Public Property Get AfectaIvaItem() As Boolean
    AfectaIvaItem = mProps.AfectaIvaItem
End Property



Public Property Let Orden(ByVal value As Integer)
    mProps.Orden = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property





Public Property Get auxSigno() As Integer
    Dim obj As IVRecargo
    
    'Si ya una vez ha recuperado (-9 significa que no está recuperado)
    If mProps.auxSigno <> -9 Then
        auxSigno = mProps.auxSigno
    Else
        Set obj = mProps.objGNComprobante.Empresa.RecuperaIVRecargo(IdRecargo)
        If Not obj Is Nothing Then
            mProps.auxSigno = obj.Signo
            mProps.auxDescripcion = obj.Descripcion
            auxSigno = mProps.auxSigno
        End If
        Set obj = Nothing
    End If
End Property

Public Property Get auxDescripcion() As String
    Dim obj As IVRecargo
    
    'Si ya una vez ha recuperado
    If Len(mProps.auxDescripcion) > 0 Then
        auxDescripcion = mProps.auxDescripcion
    Else
        Set obj = mProps.objGNComprobante.Empresa.RecuperaIVRecargo(IdRecargo)
        If Not obj Is Nothing Then
            mProps.auxSigno = obj.Signo
            mProps.auxDescripcion = obj.Descripcion
            auxDescripcion = mProps.auxDescripcion
        End If
        Set obj = Nothing
    End If
End Property

Private Sub Class_Initialize()
    mProps.auxSigno = -9        'Indica que no está asignado
End Sub

Private Sub Class_Terminate()
    Set mProps.objGNComprobante = Nothing
End Sub


