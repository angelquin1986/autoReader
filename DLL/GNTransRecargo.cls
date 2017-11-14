VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GNTransRecargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"GNTrans"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Private Type T_PROP
    Id As Long
    IdRecargo As Long
    CodRecargo As String
    ValorPre As Currency
    PorcentajePre As Currency
    BandModificable As Boolean
    BandSeleccionable As Boolean    '*** Diego 02/10/2003
    BandOrigen As Integer
    BandProrrateado As Boolean
    AfectaIvaItem As Boolean        '*** MAKOTO 13/dic/00
    Orden As Integer
    objGNTrans As GNTrans           'Referencia al objeto padre
End Type

Private mProps As T_PROP



Friend Sub Grabar()
    Dim sql As String, rs As Recordset

    'Verifica datos
    If mProps.IdRecargo = 0 And mProps.BandOrigen <> REC_SUBTOTAL Then
        Err.Raise ERR_INVALIDO, "GNTransRecargo_Grabar", MSGERR_INVALIDO
    End If


    sql = "SELECT * FROM GNTransRecargo WHERE 1=0"
    Set rs = mProps.objGNTrans.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !CodTrans = mProps.objGNTrans.CodTrans
        !IdRecargo = mProps.IdRecargo
        !ValorPre = mProps.ValorPre
        !PorcentajePre = mProps.PorcentajePre
        !BandModificable = mProps.BandModificable
        !BandSeleccionable = mProps.BandSeleccionable
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

Friend Sub Recuperar(Id As Long)
    Dim sql As String, rs As Recordset

    mProps.Id = Id
    sql = "SELECT A.*,B.CodRecargo FROM GNTransRecargo A LEFT JOIN IVRecargo B ON A.IdRecargo=B.IdRecargo " & _
          "WHERE Id=" & mProps.Id
    Set rs = mProps.objGNTrans.Empresa.OpenRecordset(sql)
    With rs
        If Not .EOF Then
            If Not IsNull(!IdRecargo) Then mProps.IdRecargo = !IdRecargo
            If Not IsNull(!CodRecargo) Then mProps.CodRecargo = !CodRecargo
            If Not IsNull(!ValorPre) Then mProps.ValorPre = !ValorPre
            If Not IsNull(!PorcentajePre) Then mProps.PorcentajePre = !PorcentajePre
            If Not IsNull(!BandModificable) Then mProps.BandModificable = !BandModificable
            If Not IsNull(!BandSeleccionable) Then mProps.BandSeleccionable = !BandSeleccionable
            If Not IsNull(!BandOrigen) Then mProps.BandOrigen = !BandOrigen
            If Not IsNull(!BandProrrateado) Then mProps.BandProrrateado = !BandProrrateado
            If Not IsNull(!AfectaIvaItem) Then mProps.AfectaIvaItem = !AfectaIvaItem    '*** MAKOTO 13/dic/00
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
        End If
        .Close
    End With
    Set rs = Nothing
End Sub

Friend Function Clone() As GNTransRecargo
    Dim nuevo As GNTransRecargo
    
    Set nuevo = New GNTransRecargo
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROP)          'Para el método Clone
    mProps = value
    mProps.Id = 0
    mProps.objGNTrans.SetModificado
End Property

Friend Property Set GNTrans(value As GNTrans)
    Set mProps.objGNTrans = value
End Property

Public Property Get GNTrans() As GNTrans
    Set GNTrans = mProps.objGNTrans
End Property





Public Property Get Id() As Long
    Id = mProps.Id
End Property


Public Property Let IdRecargo(ByVal value As Long)
    mProps.IdRecargo = value
    mProps.objGNTrans.SetModificado
End Property

Public Property Get IdRecargo() As Long
    IdRecargo = mProps.IdRecargo
End Property


Public Property Let CodRecargo(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodRecargo Then
        If Len(value) > 0 Then
            sql = "SELECT IdRecargo FROM IVRecargo WHERE CodRecargo='" & value & "'"
            Set rs = mProps.objGNTrans.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNTransRecargo_CodRecargo", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdRecargo = rs!IdRecargo
            End If
            rs.Close
        Else
            IdRecargo = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodRecargo = value
    mProps.objGNTrans.SetModificado
End Property

Public Property Get CodRecargo() As String
    CodRecargo = Trim$(mProps.CodRecargo)
End Property


Public Property Let ValorPre(ByVal value As Currency)
    mProps.ValorPre = value
    mProps.objGNTrans.SetModificado
End Property

Public Property Get ValorPre() As Currency
    ValorPre = mProps.ValorPre
End Property

Public Property Let PorcentajePre(ByVal value As Currency)
    If value < 0 Or value > 1 Then
        Err.Raise ERR_INVALIDO, "GNTransRecargo_PorcentajePre", MSGERR_INVALIDO
    End If
    
    mProps.PorcentajePre = value
    mProps.objGNTrans.SetModificado
End Property

Public Property Get PorcentajePre() As Currency
    PorcentajePre = mProps.PorcentajePre
End Property

Public Property Let BandModificable(ByVal value As Boolean)
    mProps.BandModificable = value
    mProps.objGNTrans.SetModificado
End Property

Public Property Get BandModificable() As Boolean
    BandModificable = mProps.BandModificable
End Property

Public Property Let BandSeleccionable(ByVal value As Boolean)
    mProps.BandSeleccionable = value
    mProps.objGNTrans.SetModificado
End Property

Public Property Get BandSeleccionable() As Boolean
    BandSeleccionable = mProps.BandSeleccionable
End Property

Public Property Let BandOrigen(ByVal value As Integer)
    mProps.BandOrigen = value
    mProps.objGNTrans.SetModificado
End Property

Public Property Get BandOrigen() As Integer
    BandOrigen = mProps.BandOrigen
End Property

Public Property Let BandProrrateado(ByVal value As Boolean)
    mProps.BandProrrateado = value
    mProps.objGNTrans.SetModificado
End Property

Public Property Get BandProrrateado() As Boolean
    BandProrrateado = mProps.BandProrrateado
End Property

'*** MAKOTO 13/dic/00 Agregado
Public Property Let AfectaIvaItem(ByVal value As Boolean)
    mProps.AfectaIvaItem = value
    mProps.objGNTrans.SetModificado
End Property

'*** MAKOTO 13/dic/00 Agregado
Public Property Get AfectaIvaItem() As Boolean
    AfectaIvaItem = mProps.AfectaIvaItem
End Property

Public Property Let Orden(ByVal value As Integer)
    mProps.Orden = value
    mProps.objGNTrans.SetModificado
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property




Private Sub Class_Initialize()
    'Valores predeterminados
    With mProps
        .BandOrigen = REC_SUMA
        .BandModificable = True
        .BandSeleccionable = False
        .BandProrrateado = False
        .AfectaIvaItem = False          '*** MAKOTO 13/dic/00
        .PorcentajePre = 0
        .ValorPre = 0
    End With
End Sub

Private Sub Class_Terminate()
    Set mProps.objGNTrans = Nothing
End Sub




