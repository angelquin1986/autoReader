VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RGKardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Private Type T_PROP
    id As Long
    IDComp As Long
    CodComp As String
    IdTipoDoc As Long
    CodTipoDoc As String
    RUC  As String
    NumEstab As String
    NumPunto  As String
    Secuencial As String
    FechaEmision As Date
    Autorizacion As String
    Base0 As Currency
    BaseIVA As Currency
    BaseNOIVA As Currency
    MontoICE As Currency
    MontoIVA As Currency
    Orden As Integer
    
    objGNComprobante As GNComprobante           'Referencia al objeto padre
End Type

Private mProps As T_PROP


Public Sub VerificaDatos()
    Dim i As Integer, cad As String, cadnew As String, j As Integer, posini As Integer
    'Verifica datos
    If mProps.Base0 = 0 And mProps.BaseIVA = 0 And mProps.BaseNOIVA = 0 Then
        Err.Raise ERR_INVALIDO, "RGKardex_VerificaDatos", _
            "Se intentó grabar sin valor de Base0 ni de BaseIVA en detalle de Reembolso Gastos."
    End If
    If mProps.BaseIVA <> 0 And mProps.MontoIVA = 0 Then
        Err.Raise ERR_INVALIDO, "RGKardex_VerificaDatos", _
            "Se intentó grabar con Base12 y si valor de IVA  en detalle de Reembolso Gastos."
    End If
    
    
    If mProps.IDComp = 0 Then
        Err.Raise ERR_INVALIDO, "RGKardex_VerificaDatos", _
            "Debe seleccionar un Comprobante."
    End If
    If mProps.IdTipoDoc = 0 Then
            Err.Raise ERR_INVALIDO, "RGKardex_VerificaDatos", _
            "Debe seleccionar un tipo de documento en detalle de Reembolso Gastos."
    End If
    
    If Len(mProps.RUC) = 0 Then
            Err.Raise ERR_INVALIDO, "RGKardex_VerificaDatos", _
            "Debe Digitar Numero de CI/RUC en detalle de Reembolso Gastos."
    End If
    
    If Len(mProps.NumEstab) = 0 Then
        Err.Raise ERR_INVALIDO, "RGKardex_VerificaDatos", _
            "Debe ingresar un número de establecimiento en detalle de Reembolso Gastos."
    End If
    
    If Len(mProps.NumPunto) = 0 Then
        Err.Raise ERR_INVALIDO, "RGKardex_VerificaDatos", _
            "Debe ingresar un número de punto en detalle de Reembolso Gastos."
    End If
    
    If Len(mProps.Secuencial) = 0 Then
        Err.Raise ERR_INVALIDO, "RGKardex_VerificaDatos", _
            "Debe ingresar un número de secuencia en detalle de Reembolso Gastos."
    End If
   
    If Len(mProps.Autorizacion) = 0 Then
        Err.Raise ERR_INVALIDO, "RGKardex_VerificaDatos", _
            "Debe ingresar un número de autorizacion en detalle de Reembolso Gastos."
    End If
   
    If Len(mProps.Autorizacion) <> 9 And Len(mProps.Autorizacion) <> 37 And Len(mProps.Autorizacion) > 49 Then
        Err.Raise ERR_INVALIDO, "RGKardex_VerificaDatos", _
            "El número de autorizacion es incorrecto en detalle de Reembolso Gastos."
    End If
   
   'revisa que la fecha no sea mayor a la de la reposicion
   For i = 1 To mProps.objGNComprobante.CountRGKardex
        If mProps.FechaEmision > mProps.objGNComprobante.FechaTrans Then
        Err.Raise ERR_INVALIDO, "RGKardex_VerificaDatos", _
            "La fecha del comprobante de Reembolso de Gastos no puede ser mayor a a de la Transacción"
        End If
   Next i
    
   'revisa no se ingrese dos veses documento
'''   For i = 1 To mProps.objGNComprobante.CountRGKardex
'''        For j = i To mProps.objGNComprobante.CountRGKardex
'''            If i = j Then
'''                cad = mProps.RUC & "-" & mProps.NumEstab & "-" & mProps.NumPunto & "-" & mProps.Secuencial & "-" & mProps.Autorizacion
'''                posini = 1
'''            Else
'''                cadnew = mProps.RUC & "-" & mProps.NumEstab & "-" & mProps.NumPunto & "-" & mProps.Secuencial & "-" & mProps.Autorizacion
'''                    If cad = cadnew Then
'''                        Err.Raise ERR_INVALIDO, "RGKardex_VerificaDatos", _
'''                        "Existen documentos duplicados en este Reembolso de Gastos" & Chr(13) & "Fila: " & posini & " documneto " & cad & " y" & Chr(13) & "Fila: " & j & " documento " & cadnew
'''
'''                    End If
'''            End If
'''        Next j
'''
'''   Next i
    
    
End Sub

Friend Sub Grabar()
    Dim sql As String, rs As Recordset
    Dim IdRGKardex As Long
    
    IdRGKardex = mProps.id
    VerificaDatos
    
   
    sql = "SELECT * FROM RGKardex WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !transid = mProps.objGNComprobante.transid
        !IDComp = mProps.IDComp
        !IdTipoDoc = mProps.IdTipoDoc
        !RUC = mProps.RUC
        !NumEstab = mProps.NumEstab
        !NumPunto = mProps.NumPunto
        !Secuencial = mProps.Secuencial
        !FechaEmision = mProps.FechaEmision
        !Autorizacion = mProps.Autorizacion
        !Base0 = mProps.Base0
        !BaseIVA = mProps.BaseIVA
        !BaseNOIVA = mProps.BaseNOIVA
        !MontoICE = mProps.MontoICE
        !MontoIVA = mProps.MontoIVA
        !Orden = mProps.Orden
        
        .Update
        
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.id = !id             'Guarda el Id(Autonumerico)
        .Close
        
    End With
    Set rs = Nothing
End Sub

Friend Sub Recuperar(rs As Recordset)
    With rs
        If Not .EOF Then
            mProps.id = !id
            If Not IsNull(!IDComp) Then mProps.IDComp = !IDComp
            If Not IsNull(!CodComp) Then mProps.CodComp = !CodComp
            If Not IsNull(!IdTipoDoc) Then mProps.IdTipoDoc = !IdTipoDoc
            If Not IsNull(!CodTipoDoc) Then mProps.CodTipoDoc = !CodTipoDoc
            If Not IsNull(!RUC) Then mProps.RUC = !RUC
            If Not IsNull(!NumEstab) Then mProps.NumEstab = !NumEstab
            If Not IsNull(!NumPunto) Then mProps.NumPunto = !NumPunto
            If Not IsNull(!Secuencial) Then mProps.Secuencial = !Secuencial
            If Not IsNull(!FechaEmision) Then mProps.FechaEmision = !FechaEmision
            If Not IsNull(!Autorizacion) Then mProps.Autorizacion = !Autorizacion
            If Not IsNull(!Base0) Then mProps.Base0 = !Base0
            If Not IsNull(!BaseIVA) Then mProps.BaseIVA = !BaseIVA
            If Not IsNull(!BaseNOIVA) Then mProps.BaseNOIVA = !BaseNOIVA
            If Not IsNull(!MontoICE) Then mProps.MontoICE = !MontoICE
            If Not IsNull(!MontoIVA) Then mProps.MontoIVA = !MontoIVA
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            
        End If
    End With
End Sub

Friend Function Clone() As RGKardex
    Dim nuevo As RGKardex
    
    Set nuevo = New RGKardex
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROP)          'Para el método Clone
    mProps = value
    mProps.id = 0
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Set GNComprobante(value As GNComprobante)
    Set mProps.objGNComprobante = value
End Property

Public Property Get GNComprobante() As GNComprobante
    Set GNComprobante = mProps.objGNComprobante
End Property

Public Property Get id() As Long
    id = mProps.id
End Property

Friend Property Get IDComp() As Long
    IDComp = mProps.IDComp
End Property


Public Property Let CodComp(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodComp Then
        If Len(value) > 0 Then
            'Actualiza IDComp también
            sql = "SELECT ID FROM anexo_comprobantes WHERE CodComprobante='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "RGKardex_CodComp", MSGERR_NOHAYCODIGO
            Else
                mProps.IDComp = rs!id
            End If
            rs.Close
        Else
            mProps.IDComp = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodComp = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodComp() As String
    CodComp = Trim$(mProps.CodComp)
End Property

Friend Property Get IdTipoDoc() As Long
    IdTipoDoc = mProps.IdTipoDoc
End Property


Public Property Let CodTipoDoc(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodTipoDoc Then
        If Len(value) > 0 Then
            'Actualiza IdTipoDoc también
            sql = "SELECT Id FROM anexo_tipoDocumento WHERE CodTipoDocumento='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "RGKardex_CodTipoDoc", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdTipoDoc = rs!id
            End If
            rs.Close
        Else
            IdTipoDoc = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodTipoDoc = value
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Let IdTipoDoc(ByVal value As Long)
    mProps.IdTipoDoc = value
    mProps.objGNComprobante.SetModificado
End Property


Public Property Get CodTipoDoc() As String
    CodTipoDoc = Trim$(mProps.CodTipoDoc)
End Property

Public Property Let RUC(ByVal value As String)
    If Len(value) > 13 Then
        Err.Raise ERR_INVALIDO, "PCKardex_RUC", MSGERR_INVALIDO
    End If
    
    mProps.RUC = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get RUC() As String
    RUC = Trim$(mProps.RUC)
End Property



Public Property Let NumEstab(ByVal value As String)
    If Len(value) > 3 Then
        Err.Raise ERR_INVALIDO, "RGKardex_NumEstab", MSGERR_INVALIDO
    End If
    
    mProps.NumEstab = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get NumEstab() As String
    NumEstab = Trim$(mProps.NumEstab)
End Property

Public Property Let NumPunto(ByVal value As String)
    If Len(value) > 3 Then
        Err.Raise ERR_INVALIDO, "PCKardex_NumPunto", MSGERR_INVALIDO
    End If
    
    mProps.NumPunto = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get NumPunto() As String
    NumPunto = Trim$(mProps.NumPunto)
End Property

Public Property Let Secuencial(ByVal value As String)
    If Len(value) > 9 Then
        Err.Raise ERR_INVALIDO, "RGKardex_Secuencial", MSGERR_INVALIDO
    End If
        
    mProps.Secuencial = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Secuencial() As String
    Secuencial = Trim$(mProps.Secuencial)
End Property


Public Property Let Autorizacion(ByVal value As String)
    If Len(value) > 49 Then
        Err.Raise ERR_INVALIDO, "RGKardex_Autorizacion", MSGERR_INVALIDO
    End If
    
    mProps.Autorizacion = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Autorizacion() As String
    Autorizacion = Trim$(mProps.Autorizacion)
End Property


Public Property Let FechaEmision(ByVal value As Date)
    mProps.FechaEmision = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get FechaEmision() As Date
    FechaEmision = mProps.FechaEmision
End Property

Public Property Let Base0(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "RGKardex_Base0", MSGERR_INVALIDO
    Else
        mProps.Base0 = value
    End If
    
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Base0() As Currency
    'Devuelve en Moneda ex.
    Base0 = mProps.Base0 / mProps.objGNComprobante.Cotizacion("")
End Property

Public Property Let BaseIVA(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "RGKardex_BaseIVA", MSGERR_INVALIDO
    Else
        mProps.BaseIVA = value
    End If
    
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get BaseIVA() As Currency
    'Devuelve en Moneda ex.
    BaseIVA = mProps.BaseIVA / mProps.objGNComprobante.Cotizacion("")
End Property

Public Property Let BaseNOIVA(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "RGKardex_BaseNOIVA", MSGERR_INVALIDO
    Else
        mProps.BaseNOIVA = value
    End If
    
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get BaseNOIVA() As Currency
    'Devuelve en Moneda ex.
    BaseNOIVA = mProps.BaseNOIVA / mProps.objGNComprobante.Cotizacion("")
End Property

Public Property Let MontoICE(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "RGKardex_MontoICE", MSGERR_INVALIDO
    Else
        mProps.MontoICE = value
    End If
    
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get MontoICE() As Currency
    'Devuelve en Moneda ex.
    MontoICE = mProps.MontoICE / mProps.objGNComprobante.Cotizacion("")
End Property

Public Property Let MontoIVA(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "RGKardex_MontoIVA", MSGERR_INVALIDO
    Else
        mProps.MontoIVA = value
    End If
    
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get MontoIVA() As Currency
    'Devuelve en Moneda ex.
    MontoIVA = mProps.MontoIVA / mProps.objGNComprobante.Cotizacion("")
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

