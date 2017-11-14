VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GNCompImpExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''
'''Option Explicit
'''
'''Private Type T_PROPS
'''    TransID As Long
'''    objEmpresa As Empresa
'''End Type

Option Explicit

Private Type T_PROP
    Id As Long
    ObservacionImpExp As String
    Consignatario As String
    FormaEnvio As String
    PesoNeto As String
    PesoBruto As String
    Referencia As String
    objGNComprobante As GNComprobante           'Referencia al objeto padre
End Type

Private mProps As T_PROP

Public Sub VerificaDatos()
End Sub

Friend Sub Grabar()
    Dim sql As String, rs As Recordset
    
    VerificaDatos
    
    sql = "SELECT * FROM GNCompIE WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !TransID = mProps.objGNComprobante.TransID
        !ObservacionImpExp = mProps.ObservacionImpExp
        !Consignatario = mProps.Consignatario
        !FormaEnvio = mProps.FormaEnvio
        !PesoNeto = mProps.PesoNeto
        !PesoBruto = mProps.PesoBruto
        !Referencia = mProps.Referencia
        .Update
        
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.Id = !Id             'Guarda el Id(Autonumerico)
        .Close
    End With
End Sub

Friend Sub Recuperar(rs As Recordset)
    With rs
        If Not .EOF Then
            mProps.Id = !Id
            If Not IsNull(!ObservacionImpExp) Then mProps.ObservacionImpExp = !ObservacionImpExp
            If Not IsNull(!Consignatario) Then mProps.Consignatario = !Consignatario
            If Not IsNull(!FormaEnvio) Then mProps.FormaEnvio = !FormaEnvio
            If Not IsNull(!PesoNeto) Then mProps.PesoNeto = !PesoNeto
            If Not IsNull(!PesoBruto) Then mProps.PesoBruto = !PesoBruto
            If Not IsNull(!Referencia) Then mProps.Referencia = !Referencia
        End If
    End With
End Sub

Friend Function Clone() As GNCompImpExp
    Dim nuevo As GNCompImpExp
    
    Set nuevo = New GNCompImpExp
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

Friend Property Get Id() As Long
    Id = mProps.Id
End Property

Private Sub Class_Terminate()
    Set mProps.objGNComprobante = Nothing
End Sub

Public Property Let ObservacionImpExp(ByVal value As String)
    value = RTrim$(value)
    If Len(value) > 300 Then
        Err.Raise ERR_INVALIDO, "IVKardex_ObservacionImpExp", MSGERR_INVALIDO
    End If
    
    mProps.ObservacionImpExp = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get ObservacionImpExp() As String
    ObservacionImpExp = RTrim$(mProps.ObservacionImpExp)
End Property



Public Property Let Consignatario(ByVal value As String)
    value = RTrim$(value)
    If Len(value) > 300 Then
        Err.Raise ERR_INVALIDO, "IVKardex_Consignatario", MSGERR_INVALIDO
    End If
    
    mProps.Consignatario = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Consignatario() As String
    Consignatario = RTrim$(mProps.Consignatario)
End Property

Public Property Let FormaEnvio(ByVal value As String)
    value = RTrim$(value)
    If Len(value) > 120 Then
        Err.Raise ERR_INVALIDO, "IVKardex_FormaEnvio", MSGERR_INVALIDO
    End If
    
    mProps.FormaEnvio = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get FormaEnvio() As String
    FormaEnvio = RTrim$(mProps.FormaEnvio)
End Property

Public Property Let PesoNeto(ByVal value As String)
    value = RTrim$(value)
    If Len(value) > 60 Then
        Err.Raise ERR_INVALIDO, "IVKardex_PesoNeto", MSGERR_INVALIDO
    End If
    
    mProps.PesoNeto = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get PesoNeto() As String
    PesoNeto = RTrim$(mProps.PesoNeto)
End Property

Public Property Let PesoBruto(ByVal value As String)
    value = RTrim$(value)
    If Len(value) > 60 Then
        Err.Raise ERR_INVALIDO, "IVKardex_PesoBruto", MSGERR_INVALIDO
    End If
    
    mProps.PesoBruto = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get PesoBruto() As String
    PesoBruto = RTrim$(mProps.PesoBruto)
End Property

Public Property Let Referencia(ByVal value As String)
    value = RTrim$(value)
    If Len(value) > 300 Then
        Err.Raise ERR_INVALIDO, "IVKardex_Referencia", MSGERR_INVALIDO
    End If
    
    mProps.Referencia = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Referencia() As String
    Referencia = RTrim$(mProps.Referencia)
End Property

