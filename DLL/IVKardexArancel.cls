VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVKardexArancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Type T_PROP
    Id As Long
    IdARANCEL As Long
    CodArancel As String
    Porcentaje As Currency
    FOB As Currency
    Seguro As Currency
    Flete As Currency
    CIF As Currency
    Arancel As Currency
    Orden As Integer
    objGNComprobante As gnComprobante           'Referencia al objeto padre
End Type

Private mProps As T_PROP

Public Sub VerificaDatos()
    'Verifica datos
    If mProps.IdARANCEL = 0 Then
        Err.Raise ERR_INVALIDO, "IVKardexArancel_VerificaDatos", _
            "Código de Arancel no está asignado."
    End If
End Sub

Friend Sub Grabar()
    Dim sql As String, rs As Recordset
    
    VerificaDatos
    
    '------
    'Aquí podría poner If para que no grabe si no tiene Seguro (Pendiente)
    '------
    
    sql = "SELECT * FROM IVKardexArancel WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !TransID = mProps.objGNComprobante.TransID
        !IdARANCEL = mProps.IdARANCEL
        !Porcentaje = mProps.Porcentaje
        !FOB = mProps.FOB
        !Seguro = mProps.Seguro
        !Flete = mProps.Flete
        !CIF = mProps.CIF
        !Orden = mProps.Orden
        !Arancel = mProps.Arancel
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
            If Not IsNull(!IdARANCEL) Then mProps.IdARANCEL = !IdARANCEL
            If Not IsNull(!CodArancel) Then mProps.CodArancel = !CodArancel
            mProps.Porcentaje = !Porcentaje
            mProps.FOB = !FOB
            mProps.Seguro = !Seguro
            mProps.Flete = !Flete
            mProps.CIF = !CIF
            mProps.Orden = !Orden
            mProps.Arancel = !Arancel
        End If
    End With
End Sub

Friend Function Clone() As IVKardexArancel
    Dim nuevo As IVKardexArancel
    
    Set nuevo = New IVKardexArancel
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROP)          'Para el método Clone
    mProps = value
    mProps.Id = 0
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Set gnComprobante(value As gnComprobante)
    Set mProps.objGNComprobante = value
End Property

Public Property Get gnComprobante() As gnComprobante
    Set gnComprobante = mProps.objGNComprobante
End Property


Public Property Get Id() As Long
    Id = mProps.Id
End Property

Public Property Let IdARANCEL(ByVal value As Long)
    mProps.IdARANCEL = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get IdARANCEL() As Long
    IdARANCEL = mProps.IdARANCEL
End Property


Public Property Let CodArancel(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el Seguro
    If value <> mProps.CodArancel Then
        'Actualiza IdArancel también
        sql = "SELECT IdArancel FROM IVRecargoArancel WHERE CodArancel='" & value & "'"
        Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "IVKardexArancel_CodArancel", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            IdARANCEL = rs!IdARANCEL
        End If
        rs.Close
    End If

    Set rs = Nothing
    mProps.CodArancel = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodArancel() As String
    CodArancel = Trim$(mProps.CodArancel)
End Property


Public Property Let Seguro(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    
    mProps.Seguro = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Seguro() As Currency
    'Devuelve en Moneda Ex.
    Seguro = mProps.Seguro / mProps.objGNComprobante.Cotizacion("")
End Property

Public Property Let Flete(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    
    mProps.Flete = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Flete() As Currency
    'Devuelve en Moneda Ex.
    Flete = mProps.Flete / mProps.objGNComprobante.Cotizacion("")
End Property

Public Property Let Orden(ByVal value As Integer)
    mProps.Orden = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property

Private Sub Class_Terminate()
    Set mProps.objGNComprobante = Nothing
End Sub

Public Property Let FOB(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    
    mProps.FOB = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get FOB() As Currency
    'Devuelve en Moneda Ex.
    FOB = mProps.FOB / mProps.objGNComprobante.Cotizacion("")
End Property

Public Property Let CIF(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    
    mProps.CIF = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CIF() As Currency
    'Devuelve en Moneda Ex.
    CIF = mProps.CIF / mProps.objGNComprobante.Cotizacion("")
End Property


Public Property Let Arancel(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    
    mProps.Arancel = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Arancel() As Currency
    'Devuelve en Moneda Ex.
    Arancel = mProps.Arancel / mProps.objGNComprobante.Cotizacion("")
End Property


