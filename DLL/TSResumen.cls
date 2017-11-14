VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TSResumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'AUC Resumen cierres
Option Explicit

Private Type T_PROP
    Id As Long
    TransID As Long
    IdDoc As Long
    ValorSis As Currency
    ValorDoc As Currency
    Diferencia As Currency
    Orden As Integer
    FechaGrabado As Date
    objGNComprobante As gnComprobante           'Referencia al objeto padre
End Type

Private mProps As T_PROP


Friend Sub Grabar()
    Dim sql As String, rs As Recordset
    
    'VerificaDatos
   
    sql = "SELECT * FROM TSResumen WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !TransID = mProps.objGNComprobante.TransID
        !IdDoc = mProps.IdDoc
        !ValorSis = mProps.ValorSis
        !ValorDoc = mProps.ValorDoc
        !Diferencia = mProps.Diferencia
        !Orden = mProps.Orden
        !FechaGrabado = mProps.objGNComprobante.FechaTrans
        
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
    With rs
        If Not .EOF Then
            mProps.Id = !Id
            If Not IsNull(!ValorSis) Then mProps.ValorSis = !ValorSis
            If Not IsNull(!ValorDoc) Then mProps.ValorDoc = !ValorDoc
            If Not IsNull(!IdDoc) Then mProps.IdDoc = !IdDoc
            If Not IsNull(!Diferencia) Then mProps.Diferencia = !Diferencia
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            If Not IsNull(!FechaGrabado) Then mProps.FechaGrabado = !FechaGrabado
        End If
    End With
End Sub

Friend Function Clone() As TSResumen
    Dim nuevo As TSResumen
    
    Set nuevo = New TSResumen
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



Public Property Let FechaGrabado(ByVal value As Date)
    mProps.FechaGrabado = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get FechaGrabado() As Date
    FechaGrabado = mProps.FechaGrabado
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

Public Property Let ValorSis(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "TSResumen_ValorSis", MSGERR_INVALIDO
    ElseIf value > 0 Then
        mProps.ValorSis = value
    Else
        mProps.ValorSis = 0
    End If
    
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get ValorSis() As Currency
    'Devuelve en Moneda ex.
    ValorSis = mProps.ValorSis / mProps.objGNComprobante.Cotizacion("")
End Property



Public Property Let ValorDoc(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "TSResumen_ValorDoc", MSGERR_INVALIDO
    ElseIf value > 0 Then
        mProps.ValorDoc = value
    Else
        mProps.ValorDoc = 0
    End If
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get ValorDoc() As Currency
    'Devuelve en Moneda ex.
    ValorDoc = mProps.ValorDoc / mProps.objGNComprobante.Cotizacion("")
End Property


Public Property Let IdDoc(ByVal value As Long)
    mProps.IdDoc = value
    mProps.objGNComprobante.SetModificado
End Property


Public Property Get IdDoc() As Long
    IdDoc = mProps.IdDoc
End Property


Public Property Let Diferencia(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
'    If value < 0 Then
 '       Err.Raise ERR_INVALIDO, "TSResumen_Diferencia", MSGERR_INVALIDO
 '   ElseIf value > 0 Then
        mProps.Diferencia = value
  '  Else
   '     mProps.Diferencia = 0
   ' End If
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Diferencia() As Currency
    'Devuelve en Moneda ex.
    Diferencia = mProps.Diferencia / mProps.objGNComprobante.Cotizacion("")
End Property

