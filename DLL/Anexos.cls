VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Anexos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Empresa"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'AUC 27/03/06
Option Explicit
Private Type T_PROPS
    IdComprobante As Long
    CodComprobante As String
    Descripcion As String
    SecuencialValidos As String
    Sustentos As String
    bandValidaComp As Boolean
    BandAutoImpresor As Boolean
    BandAutorizaActiva As Boolean
    
    IdSustento As Long
    CodSustento As String
    DescripcionSust As String
    ComprobantesValidos As String
    bandValidaSust As Boolean
    
    IdRetencion As Long
    CodRetencion As String
    DescripcionRetIR As String
    Porcentaje As Currency
    bandValidaRet As Boolean
    
    IdTrans As Long
    CodTrans As String
    DescripcionAnexoTrans As String
    ComprobantesValidosTrans As String
    SecuencialTransValidos As String
    bandValidaTrans As Boolean
   
   ' jeaa 15/05/2007 Tipo Documento
    IdTipoDocumento As Long
    codtipoDocumento As String
    DescripcionTipoDocumento As String
    bandValidaTipoDocumento As Boolean
   
   ' jeaa 15/05/2007 Tipo Documento
    IdDistrito As Long
    CodDistrito As String
    DescripcionDistrito As String
    bandValidaDistrito As Boolean
   
   ' jeaa 15/05/2007 Tipo Documento
    IdRegimen As Long
    CodRegimen As String
    DescripcionRegimen As String
    bandValidaRegimen As Boolean
   
   ' jeaa 15/05/2007 Tipo Documento
    IdTipoEXp As Long
    CodTipoExp As String
    DescripcionTipoExp As String
    bandValidaTipoExp As Boolean
   
   
   ' jeaa 15/05/2007 Tipo FomaPago
    IdFormaPago As Long
    CodFormaPago As String
    DescripcionFormaPago As String
    bandValidaFormaPago As Boolean
    
    'jeaa 15/agt/ 2013 pais
   
    IdPais As Long
    CodPais As String
    DescripcionPais As String
    bandValidaPais As Boolean
    
    'JEAA 23/07/2016
   
    IdICE As Long
    CodICE As String
    DescripcionICE As String
    bandValidaICE As Boolean
    CodClasifica As String
   
   
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Public Sub Grabar()
    Dim sql As String, rs As Recordset, rsaux As Recordset, rsaux1 As Recordset
    Dim resp As Integer, i As Integer
    
        'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si este tipo de transaccion esta configurado como autoimpresor
    
    
    'Verifica si no repite el código
    sql = "SELECT * FROM Anexo_Comprobantes" & _
          " WHERE  codComprobante  ='" & mProps.CodComprobante & "'" & _
          "And Id =  " & mProps.IdComprobante
          
   
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM  Anexo_comprobantes  WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        'ojo para poder cambiar un codigo se debe crear un id en c/tabla
        sql = "SELECT * FROM Anexo_Comprobantes WHERE id =" & mProps.IdComprobante
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
End If
    'Graba el registro en Pos
    With rs
        !CodComprobante = mProps.CodComprobante
        !Descripcion = mProps.Descripcion
        !SecuencialValidos = mProps.SecuencialValidos
        !Sustentos = mProps.Sustentos
        !BandValida = mProps.bandValidaComp
''''        If !BandAutoImpresor And Not mProps.BandAutoImpresor Then
''''            sql = "SELECT codtrans, opcion FROM gntrans"
''''            sql = sql & " WHERE  AnexoCodTipoComp  ='" & mProps.CodComprobante & "'"
''''            sql = sql & " and substring(opcion,152,1) ='S'"
''''            Set rsaux = mProps.objEmpresa.OpenRecordset(sql)
''''            resp = vbCancel
''''            If rs.RecordCount > 0 Then
''''                resp = MsgBox("Éste tipo de comprobante utiliza autoimpresor," & Chr(13) & "si deshabilita la opción de autoimpresor," & Chr(13) & " se desactivaran las transacciones asociadas", vbOKCancel)
''''                If resp = vbCancel Then
''''                    Exit Sub
''''                Else
''''                    For i = 1 To rsaux.RecordCount
''''                        sql = " Update Gntrans "
''''                        sql = sql & " set opcion ='" & Mid$(rsaux.Fields("opcion"), 1, 151) & "N" & Mid$(rsaux.Fields("opcion"), 153, 50) & "', "
''''                        sql = sql & " BandValida= 0 "
''''                        sql = sql & " where codtrans='" & rsaux.Fields("codtrans") & "'"
''''                        Set rsaux1 = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
''''                        rsaux.MoveNext
''''                    Next i
''''
''''                End If
''''            End If
''''        End If
        
        
        !BandAutoImpresor = mProps.BandAutoImpresor
        !BandAutorizaActiva = mProps.BandAutorizaActiva
        .Update
        
        
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.IdComprobante = !id
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_ANCOMP", "AnexoComp Nuevo: " & _
                                                         mProps.CodComprobante, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_ANCOMP", "AnexoComp Modificado: " & _
                                                         mProps.CodComprobante, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_ANEXOCOMPROBANTES
End Sub

Friend Sub RecuperarAnexoComprobante(key As Variant)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT * FROM  Anexo_Comprobantes WHERE   codcomprobante ='" & key & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        mProps.IdComprobante = !id
        mProps.CodComprobante = !CodComprobante
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!SecuencialValidos) Then mProps.SecuencialValidos = !SecuencialValidos
        If Not IsNull(!Sustentos) Then mProps.Sustentos = !Sustentos
        If Not IsNull(!BandValida) Then mProps.bandValidaComp = !BandValida
        If Not IsNull(!BandAutoImpresor) Then mProps.BandAutoImpresor = !BandAutoImpresor
        If Not IsNull(!BandAutorizaActiva) Then mProps.BandAutorizaActiva = !BandAutorizaActiva
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String
    Dim rs As Recordset
    

    sql = "SELECT CodComprobante FROM Anexo_Comprobantes" & _
          " WHERE codComprobante= '" & mProps.CodComprobante & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
'    If Not rs.EOF Then
'        Err.Raise ERR_NOELIMINABLE, "Anexos_Eliminar", MSGERR_NOELIMINABLE
'    End If
'   rs.Close

    'Elimina el registro
    sql = "DELETE FROM Anexo_Comprobantes WHERE codComprobante= '" & mProps.CodComprobante & "'"
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_ANCOMP", "AnexoComprobante Eliminado: " & _
                                                         mProps.CodComprobante, "DT"
                                                         
    Set rs = Nothing
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_ANEXOCOMPROBANTES
End Sub

Friend Function Clone() As Anexos
    Dim nuevo As Anexos
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New Anexos
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdComprobante = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property




'Friend Property Let NumGrupo(ByVal value As Byte)
'    mProps.NumGrupo = value
'    mbooModificado = True
'End Property
'
'Public Property Get NumGrupo() As Byte
'    NumGrupo = mProps.NumGrupo
'End Property


Public Property Let CodComprobante(ByVal value As String)
    ValidaCodigo value      'Valida si contiene caracteres no validos
    mProps.CodComprobante = value
    mbooModificado = True
End Property

Public Property Get CodComprobante() As String
    CodComprobante = Trim$(mProps.CodComprobante)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 120 Then
        Err.Raise ERR_INVALIDO, "Anexos_Descripcion", MSGERR_INVALIDO
    End If
    mProps.Descripcion = value
    mbooModificado = True
End Property
Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property
Public Property Let SecuencialValidos(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "Anexos_SecuencialValidos", MSGERR_INVALIDO
    End If
    mProps.SecuencialValidos = value
    mbooModificado = True
End Property
Public Property Get SecuencialValidos() As String
    SecuencialValidos = Trim$(mProps.SecuencialValidos)
End Property
Public Property Let Sustentos(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "Anexos_Sustentos", MSGERR_INVALIDO
    End If
    mProps.Sustentos = value
    mbooModificado = True
End Property
Public Property Get Sustentos() As String
    Sustentos = Trim$(mProps.Sustentos)
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
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub


Public Property Let CodSustento(ByVal value As String)
    If Len(value) > 2 Then
        Err.Raise ERR_INVALIDO, "Anexos_DescripcionSust", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    mProps.CodSustento = value
    mbooModificado = True
End Property
Public Property Get CodSustento() As String
    CodSustento = Trim$(mProps.CodSustento)
End Property
Public Property Let DescripcionSust(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "Anexos_DescripcionSust", MSGERR_INVALIDO
    End If
    mProps.DescripcionSust = value
    mbooModificado = True
End Property
Public Property Get DescripcionSust() As String
    DescripcionSust = Trim$(mProps.DescripcionSust)
End Property
Public Property Let ComprobantesValidos(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Anexos_ComprobantesValidos", MSGERR_INVALIDO
    End If
    mProps.ComprobantesValidos = value
    mbooModificado = True
End Property
Public Property Get ComprobantesValidos() As String
    ComprobantesValidos = Trim$(mProps.ComprobantesValidos)
End Property

Friend Sub RecuperarAnexoSust(key As Variant)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT * FROM  Anexo_suntentos WHERE  codsustento ='" & key & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        mProps.IdSustento = !id
        mProps.CodSustento = !CodSustento
        If Not IsNull(!Descripcion) Then mProps.DescripcionSust = !Descripcion
        If Not IsNull(!BandValida) Then mProps.bandValidaSust = !BandValida
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub GrabarSust()
    Dim sql As String, rs As Recordset
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
    sql = "SELECT * FROM Anexo_suntentos" & _
          " WHERE  codsustento  ='" & mProps.CodSustento & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM  Anexo_suntentos WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM Anexo_suntentos WHERE id ='" & mProps.IdSustento & "'"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
    End If

    'Graba el registro en Pos
    With rs
        !CodSustento = mProps.CodSustento
        !Descripcion = mProps.DescripcionSust
        !ComprobantesValidos = mProps.ComprobantesValidos
        !BandValida = mProps.bandValidaSust
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
         mProps.IdSustento = !id
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_ANSUST", "AnexoSust Nuevo: " & _
                                                         mProps.CodSustento, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_ANSUST", "AnexoSust Modificado: " & _
                                                         mProps.CodSustento, "DT"
    End If
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_ANEXOSUSTENTOS
End Sub


Public Sub EliminarSust()
    Dim sql As String
    Dim rs As Recordset

    sql = "SELECT Codsustento FROM Anexo_suntentos" & _
          " WHERE codsustento= '" & mProps.CodSustento & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
'    If Not rs.EOF Then
'        Err.Raise ERR_NOELIMINABLE, "Anexos_Eliminar", MSGERR_NOELIMINABLE
'    End If
'   rs.Close

    'Elimina el registro
    sql = "DELETE FROM Anexo_suntentos WHERE codsustento= '" & mProps.CodSustento & "'"
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_ANSUST", "AnexoSust Eliminado: " & _
                                                         mProps.CodSustento, "DT"
                                                         
    Set rs = Nothing
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_ANEXOSUSTENTOS
End Sub

Public Property Let CodRetencion(ByVal value As String)
    If Len(value) > 5 Then
        Err.Raise ERR_INVALIDO, "Anexos_CodigoRet", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    mProps.CodRetencion = value
    mbooModificado = True
End Property
Public Property Get CodRetencion() As String
    CodRetencion = Trim$(mProps.CodRetencion)
End Property
Public Property Let DescripcionRetIR(ByVal value As String)
    If Len(value) > 180 Then
        Err.Raise ERR_INVALIDO, "Anexos_DescripcionRetIR", MSGERR_INVALIDO
    End If
    mProps.DescripcionRetIR = value
    mbooModificado = True
End Property
Public Property Get DescripcionRetIR() As String
    DescripcionRetIR = Trim$(mProps.DescripcionRetIR)
End Property
Public Property Let Porcentaje(ByVal value As Currency)
    mProps.Porcentaje = value
    mbooModificado = True
End Property
Public Property Get Porcentaje() As Currency
    Porcentaje = mProps.Porcentaje
End Property

Friend Sub RecuperarAnexoRetIR(key As Variant)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT * FROM  Anexos_retencion_ir WHERE  codretencion ='" & key & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        mProps.IdRetencion = !id
        mProps.CodRetencion = !CodRetencion
        If Not IsNull(!Descripcion) Then mProps.DescripcionRetIR = !Descripcion
        If Not IsNull(!Porcentaje) Then mProps.Porcentaje = !Porcentaje
        If Not IsNull(!BandValida) Then mProps.bandValidaRet = !BandValida
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub GrabarRetIR()
    Dim sql As String, rs As Recordset
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
    sql = "SELECT * FROM Anexos_Retencion_IR" & _
          " WHERE  codretencion  ='" & mProps.CodRetencion & "'" & _
          " AND Id = " & mProps.IdRetencion
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM  Anexos_Retencion_IR WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM Anexos_Retencion_IR WHERE id ='" & mProps.IdRetencion & "'"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
    End If

    'Graba el registro en Pos
    With rs
        '!Id = mProps.IdRetencion
        !CodRetencion = mProps.CodRetencion
        !Descripcion = mProps.DescripcionRetIR
        !Porcentaje = mProps.Porcentaje
        !BandValida = mProps.bandValidaRet
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
         mProps.IdRetencion = !id
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_ANRTIR", "AnexoRETIR Nuevo: " & _
                                                         mProps.CodRetencion, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_ANRTIR", "AnexoRETIR Modificado: " & _
                                                         mProps.CodRetencion, "DT"
    End If
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_ANEXORETIR
End Sub

Public Sub EliminarRetIR()
    Dim sql As String
    Dim rs As Recordset

    sql = "SELECT CodRetencion FROM Anexos_Retencion_IR" & _
          " WHERE codRetencion= '" & mProps.CodRetencion & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
'    If Not rs.EOF Then
'        Err.Raise ERR_NOELIMINABLE, "Anexos_Eliminar", MSGERR_NOELIMINABLE
'    End If
'   rs.Close
    'Elimina el registro
    sql = "DELETE FROM Anexos_Retencion_IR WHERE codRetencion= '" & mProps.CodRetencion & "'"
    mProps.objEmpresa.Execute sql, True
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_ANRETIR", "AnexoRETIR Eliminado: " & _
                                                         mProps.CodRetencion, "DT"
    Set rs = Nothing
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_ANEXORETIR
End Sub

Public Property Let CodTrans(ByVal value As String)
    If Len(value) > 5 Then
        Err.Raise ERR_INVALIDO, "Anexos_CodigoTrans", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    mProps.CodTrans = value
    mbooModificado = True
End Property
Public Property Get CodTrans() As String
    CodTrans = Trim$(mProps.CodTrans)
End Property
Public Property Let DescripcionAnexosTrans(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "Anexos_DescripcionAnexosTrans", MSGERR_INVALIDO
    End If
    mProps.DescripcionAnexoTrans = value
    mbooModificado = True
End Property
Public Property Get DescripcionAnexosTrans() As String
    DescripcionAnexosTrans = Trim$(mProps.DescripcionAnexoTrans)
End Property
Public Property Let ComprobantesValidosTrans(ByVal value As String)
    If Len(value) > 50 Then
          Err.Raise ERR_INVALIDO, "Anexos_ComprobantesValidosTrans", MSGERR_INVALIDO
    End If
    mProps.ComprobantesValidosTrans = value
    mbooModificado = True
End Property
Public Property Get ComprobantesValidosTrans() As String
    ComprobantesValidosTrans = Trim$(mProps.ComprobantesValidosTrans)
End Property
Public Property Let SecuencialTransValidos(ByVal value As String)
    If Len(value) > 50 Then
          Err.Raise ERR_INVALIDO, "Anexos_SecuencialTransValidos", MSGERR_INVALIDO
    End If
    mProps.SecuencialTransValidos = value
    mbooModificado = True
End Property
Public Property Get SecuencialTransValidos() As String
    SecuencialTransValidos = Trim$(mProps.SecuencialTransValidos)
End Property

Friend Sub RecuperarAnexosTrans(key As Variant)
    Dim sql As String, rs As Recordset
        
    sql = "SELECT * FROM  Anexos_Transacciones WHERE  codtrans ='" & key & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        mProps.IdTrans = !id
        mProps.CodTrans = !CodTrans
        If Not IsNull(!Descripcion) Then mProps.DescripcionAnexoTrans = !Descripcion
        If Not IsNull(!ComprobantesValidos) Then mProps.ComprobantesValidos = !ComprobantesValidos
        If Not IsNull(!SecuencialTransValidos) Then mProps.SecuencialTransValidos = !SecuencialTransValidos
        If Not IsNull(!BandValida) Then mProps.bandValidaTrans = !BandValida
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub GrabarAnexosTrans()
    Dim sql As String, rs As Recordset
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
    sql = "SELECT * FROM Anexos_Transacciones " & _
          " WHERE  codTrans  ='" & mProps.CodTrans & "'" & _
          " AND id =" & mProps.IdTrans
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM  Anexos_Transacciones WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM Anexos_Transacciones WHERE id ='" & mProps.IdTrans & "'"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
    End If
    'Graba el registro en Pos
    With rs
        !CodTrans = mProps.CodTrans
        !Descripcion = mProps.DescripcionAnexoTrans
        !ComprobantesValidos = mProps.ComprobantesValidos
        !SecuencialTransValidos = mProps.SecuencialTransValidos
        !BandValida = mProps.bandValidaTrans
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.IdTrans = !id
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_ANTRNS", "AnexoTrans Nuevo: " & _
                                                         mProps.CodTrans, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_ANTRNS", "AnexoTrans Modificado: " & _
                                                         mProps.CodTrans, "DT"
    End If
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_ANEXOTRANS
End Sub

Public Sub EliminarTrans()
    Dim sql As String
    Dim rs As Recordset

    sql = "SELECT CodTrans FROM Anexos_Transacciones" & _
          " WHERE codTrans= '" & mProps.CodTrans & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
'    If Not rs.EOF Then
'        Err.Raise ERR_NOELIMINABLE, "Anexos_Eliminar", MSGERR_NOELIMINABLE
'    End If
'   rs.Close
    'Elimina el registro
    sql = "DELETE FROM Anexos_Transacciones WHERE codTrans= '" & mProps.CodTrans & "'"
    mProps.objEmpresa.Execute sql, True
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_ANTRANS", "AnexoTrans Eliminado: " & _
                                                         mProps.CodTrans, "DT"
    Set rs = Nothing
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_ANEXOTRANS
End Sub

Friend Property Get IdComprobante() As Long
    IdComprobante = mProps.IdComprobante
End Property
Friend Property Get IdSustento() As Long
    IdSustento = mProps.IdSustento
End Property
Friend Property Get IdRetencion() As Long
    IdRetencion = mProps.IdRetencion
End Property
Friend Property Get IdTrans() As Long
    IdTrans = mProps.IdTrans
End Property

Public Function ListaSustentos() As Variant
    Dim i As Long, s As String, v() As Variant, cnt As Integer
    Dim sql As String, rs As Recordset
            
            sql = "SELECT Codsustento, Descripcion FROM anexo_suntentos"
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.Open sql, gobjMain.EmpresaActual.Coneccion, adOpenStatic, adLockReadOnly
            Do While Not rs.EOF
                ReDim Preserve v(1, cnt)
                v(0, cnt) = rs!CodSustento
                v(1, cnt) = rs!Descripcion
                cnt = cnt + 1
                rs.MoveNext
            Loop
            rs.Close
        ListaSustentos = v
        Set rs = Nothing
End Function

Public Property Let bandValidaComp(ByVal value As Boolean)
    mProps.bandValidaComp = value
    mbooModificado = True
End Property

Public Property Get bandValidaComp() As Boolean
    bandValidaComp = mProps.bandValidaComp
End Property

Public Property Let bandValidaSust(ByVal value As Boolean)
    mProps.bandValidaSust = value
    mbooModificado = True
End Property

Public Property Get bandValidaSust() As Boolean
    bandValidaSust = mProps.bandValidaSust
End Property

Public Property Let bandValidaRet(ByVal value As Boolean)
    mProps.bandValidaRet = value
    mbooModificado = True
End Property

Public Property Get bandValidaRet() As Boolean
    bandValidaRet = mProps.bandValidaRet
End Property

Public Property Let bandValidaTrans(ByVal value As Boolean)
    mProps.bandValidaTrans = value
    mbooModificado = True
End Property

Public Property Get bandValidaTrans() As Boolean
    bandValidaTrans = mProps.bandValidaTrans
End Property

Public Function ListaComprobantes() As Variant
    Dim i As Long, s As String, v() As Variant, cnt As Integer
    Dim sql As String, rs As Recordset
            
            sql = "SELECT CodComprobante, Descripcion FROM anexo_comprobantes order by codcomprobante"
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.Open sql, gobjMain.EmpresaActual.Coneccion, adOpenStatic, adLockReadOnly

            Do While Not rs.EOF
                ReDim Preserve v(1, cnt)
                v(0, cnt) = rs!CodComprobante
                v(1, cnt) = rs!Descripcion
                cnt = cnt + 1
                rs.MoveNext
            Loop
            rs.Close
        ListaComprobantes = v
        Set rs = Nothing
End Function



Public Sub GrabarTipoDocumento()
    Dim sql As String, rs As Recordset
    
        'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
    sql = "SELECT * FROM Anexo_TipoDocumento" & _
          " WHERE  CodTipoDocumento  ='" & mProps.codtipoDocumento & "'" & _
          "And Id =  " & mProps.IdTipoDocumento
          
   
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM  Anexo_TipoDocumento  WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        'ojo para poder cambiar un codigo se debe crear un id en c/tabla
        sql = "SELECT * FROM Anexo_TipoDocumento WHERE id =" & mProps.IdTipoDocumento
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
End If
    'Graba el registro en Pos
    With rs
        !codtipoDocumento = mProps.codtipoDocumento
        !Descripcion = mProps.DescripcionTipoDocumento
        !BandValida = mProps.bandValidaTipoDocumento
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.IdTipoDocumento = !id
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_ANCOMP", "AnexoComp Nuevo: " & _
                                                         mProps.CodComprobante, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_ANCOMP", "AnexoComp Modificado: " & _
                                                         mProps.CodComprobante, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_ANEXOTIPODOC
End Sub

Friend Sub RecuperarAnexoTipoDocumento(key As Variant)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT * FROM  Anexo_TipoDocumento WHERE CodTipoDocumento ='" & key & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        mProps.IdTipoDocumento = !id
        mProps.codtipoDocumento = !codtipoDocumento
        If Not IsNull(!Descripcion) Then mProps.DescripcionTipoDocumento = !Descripcion
        If Not IsNull(!BandValida) Then mProps.bandValidaTipoDocumento = !BandValida
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub EliminarTipoDocumento()
    Dim sql As String
    Dim rs As Recordset
    

    'Verifica si ya existe un item usando el mismo tIPO dOCUMENTO
'    sql = "SELECT CodProvCli FROM PCProvCli " & _
'          " WHERE IDtIPOdOCUMENTO = " & mProps.IdTipoDocumento
'    Set rs = mProps.objEmpresa.OpenRecordset(sql)
'    If Not rs.EOF Then
'        Err.Raise ERR_NOELIMINABLE, "Anexo_TipoDocumento", MSGERR_NOELIMINABLE
'    End If
'    rs.Close

    'Elimina el registro
    sql = "DELETE FROM Anexo_TipoDocumento WHERE CodTipoDocumento= '" & mProps.codtipoDocumento & "'"
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_TIDOC", "AnexoTipoDocumento Eliminado: " & _
                                                         mProps.codtipoDocumento, "DT"
                                                         
    Set rs = Nothing
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_ANEXOTIPODOC
End Sub

Public Property Let codtipoDocumento(ByVal value As String)
    ValidaCodigo value      'Valida si contiene caracteres no validos
    mProps.codtipoDocumento = value
    mbooModificado = True
End Property

Public Property Get codtipoDocumento() As String
    codtipoDocumento = Trim$(mProps.codtipoDocumento)
End Property

Public Property Let DescripcionTipoDocumento(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Anexos_DescripcionTipoDocumento", MSGERR_INVALIDO
    End If
    mProps.DescripcionTipoDocumento = value
    mbooModificado = True
End Property

Public Property Get DescripcionTipoDocumento() As String
    DescripcionTipoDocumento = Trim$(mProps.DescripcionTipoDocumento)
End Property

Public Property Let bandValidaTipoDocumento(ByVal value As Boolean)
    mProps.bandValidaTipoDocumento = value
    mbooModificado = True
End Property

Public Property Get bandValidaTipoDocumento() As Boolean
    bandValidaTipoDocumento = mProps.bandValidaTipoDocumento
End Property

Friend Sub RecuperarAnexosDistrito(key As Variant)
    Dim sql As String, rs As Recordset
    sql = "SELECT * FROM  Anexo_Distrito WHERE  codDistrito ='" & key & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        mProps.IdDistrito = !IdDistrito
        mProps.CodDistrito = !CodDistrito
        If Not IsNull(!Descripcion) Then mProps.DescripcionDistrito = !Descripcion
        If Not IsNull(!BandValida) Then mProps.bandValidaDistrito = !BandValida
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub
Friend Sub RecuperarAnexosRegimen(key As Variant)
    Dim sql As String, rs As Recordset
    sql = "SELECT * FROM  Anexo_Regimen WHERE  codRegimen ='" & key & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        mProps.IdRegimen = !IdRegimen
        mProps.CodRegimen = !CodRegimen
        If Not IsNull(!Descripcion) Then mProps.DescripcionRegimen = !Descripcion
        If Not IsNull(!BandValida) Then mProps.bandValidaRegimen = !BandValida
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub
Friend Sub RecuperarAnexosTipoExp(key As Variant)
    Dim sql As String, rs As Recordset
    sql = "SELECT * FROM  Anexo_TipoExportacion WHERE  CodTipoExportacion ='" & key & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        mProps.IdTipoEXp = !Idtipoexportacion
        mProps.CodTipoExp = !CodTipoExportacion
        If Not IsNull(!Descripcion) Then mProps.DescripcionTipoExp = !Descripcion
        If Not IsNull(!BandValida) Then mProps.bandValidaTipoExp = !BandValida
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub
Public Sub GrabarDistrito()
    Dim sql As String, rs As Recordset
        'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    'Verifica si no repite el código
    sql = "SELECT * FROM Anexo_Distrito" & _
          " WHERE  CodDistrito  ='" & mProps.CodDistrito & "'" & _
          "And IdDistrito =  " & mProps.IdDistrito
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM  Anexo_Distrito  WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        'ojo para poder cambiar un codigo se debe crear un id en c/tabla
        sql = "SELECT * FROM Anexo_Distrito WHERE idDistrito =" & mProps.IdDistrito
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
End If
    'Graba el registro en Pos
    With rs
        !CodDistrito = mProps.CodDistrito
        !Descripcion = mProps.DescripcionDistrito
        !BandValida = mProps.bandValidaDistrito
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.IdDistrito = !IdDistrito
        .Close
    End With
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_ANCOMP", "AnexoComp Nuevo: " & _
                                                         mProps.CodComprobante, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_ANCOMP", "AnexoComp Modificado: " & _
                                                         mProps.CodComprobante, "DT"
    End If
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_ANEXODISTRITO
End Sub
Public Property Let CodDistrito(ByVal value As String)
    ValidaCodigo value      'Valida si contiene caracteres no validos
    mProps.CodDistrito = value
    mbooModificado = True
End Property
Public Property Get CodDistrito() As String
    CodDistrito = Trim$(mProps.CodDistrito)
End Property
Public Property Let DescripcionDistrito(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Anexos_DescripcionDistrito", MSGERR_INVALIDO
    End If
    mProps.DescripcionDistrito = value
    mbooModificado = True
End Property
Public Property Get DescripcionDistrito() As String
    DescripcionDistrito = Trim$(mProps.DescripcionDistrito)
End Property
Public Property Let bandValidaDistrito(ByVal value As Boolean)
    mProps.bandValidaDistrito = value
    mbooModificado = True
End Property
Public Property Get bandValidaDistrito() As Boolean
    bandValidaDistrito = mProps.bandValidaDistrito
End Property
Public Sub EliminarDistrito()
    Dim sql As String
    Dim rs As Recordset
    'Verifica si ya existe un item usando el mismo tIPO dOCUMENTO
'    sql = "SELECT CodProvCli FROM PCProvCli " & _
'          " WHERE IDDistrito = " & mProps.IdDistrito
'    Set rs = mProps.objEmpresa.OpenRecordset(sql)
'    If Not rs.EOF Then
'        Err.Raise ERR_NOELIMINABLE, "Anexo_Distrito", MSGERR_NOELIMINABLE
'    End If
'    rs.Close
    'Elimina el registro
    sql = "DELETE FROM Anexo_Distrito WHERE CodDistrito= '" & mProps.CodDistrito & "'"
    mProps.objEmpresa.Execute sql, True
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_TIDOC", "AnexoDistrito Eliminado: " & _
                                                         mProps.CodDistrito, "DT"
    Set rs = Nothing
    mbooEsNuevo = True
    mbooModificado = True
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_ANEXODISTRITO
End Sub
Public Sub GrabarRegimen()
    Dim sql As String, rs As Recordset
        'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    'Verifica si no repite el código
    sql = "SELECT * FROM Anexo_Regimen" & _
          " WHERE  CodRegimen  ='" & mProps.CodRegimen & "'" & _
          "And Idregimen =  " & mProps.IdRegimen
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM  Anexo_Regimen  WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        'ojo para poder cambiar un codigo se debe crear un id en c/tabla
        sql = "SELECT * FROM Anexo_Regimen WHERE idRegimen =" & mProps.IdRegimen
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
End If
    'Graba el registro en Pos
    With rs
        !CodRegimen = mProps.CodRegimen
        !Descripcion = mProps.DescripcionRegimen
        !BandValida = mProps.bandValidaRegimen
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.IdRegimen = !IdRegimen
        .Close
    End With
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_ANCOMP", "AnexoComp Nuevo: " & _
                                                         mProps.CodComprobante, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_ANCOMP", "AnexoComp Modificado: " & _
                                                         mProps.CodComprobante, "DT"
    End If
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_ANEXOREGIMEN
End Sub
Public Property Let CodRegimen(ByVal value As String)
    ValidaCodigo value      'Valida si contiene caracteres no validos
    mProps.CodRegimen = value
    mbooModificado = True
End Property
Public Property Get CodRegimen() As String
    CodRegimen = Trim$(mProps.CodRegimen)
End Property
Public Property Let DescripcionRegimen(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Anexos_DescripcionRegimen", MSGERR_INVALIDO
    End If
    mProps.DescripcionRegimen = value
    mbooModificado = True
End Property
Public Property Get DescripcionRegimen() As String
    DescripcionRegimen = Trim$(mProps.DescripcionRegimen)
End Property
Public Property Let bandValidaRegimen(ByVal value As Boolean)
    mProps.bandValidaRegimen = value
    mbooModificado = True
End Property
Public Property Get bandValidaRegimen() As Boolean
    bandValidaRegimen = mProps.bandValidaRegimen
End Property
Public Sub EliminarRegimen()
    Dim sql As String
    Dim rs As Recordset
    'Verifica si ya existe un item usando el mismo tIPO dOCUMENTO
'    sql = "SELECT CodProvCli FROM PCProvCli " & _
'          " WHERE IDRegimen = " & mProps.IdRegimen
'    Set rs = mProps.objEmpresa.OpenRecordset(sql)
'    If Not rs.EOF Then
'        Err.Raise ERR_NOELIMINABLE, "Anexo_Regimen", MSGERR_NOELIMINABLE
'    End If
'    rs.Close
    'Elimina el registro
    sql = "DELETE FROM Anexo_Regimen WHERE CodRegimen= '" & mProps.CodRegimen & "'"
    mProps.objEmpresa.Execute sql, True
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_TIDOC", "AnexoRegimen Eliminado: " & _
                                                         mProps.CodRegimen, "DT"
    Set rs = Nothing
    mbooEsNuevo = True
    mbooModificado = True
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_ANEXOREGIMEN
End Sub
Public Sub GrabarTipoExp()
    Dim sql As String, rs As Recordset
        'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    'Verifica si no repite el código
    sql = "SELECT * FROM Anexo_TipoExportacion " & _
          " WHERE  CodTipoExportacion  ='" & mProps.CodTipoExp & "'" & _
          "And IdTipoExportacion=  " & mProps.IdTipoEXp
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM  Anexo_TipoExportacion WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        'ojo para poder cambiar un codigo se debe crear un id en c/tabla
        sql = "SELECT * FROM Anexo_TipoExportacion WHERE IdTipoExportacion =" & mProps.IdTipoEXp
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
End If
    'Graba el registro en Pos
    With rs
        !CodTipoExportacion = mProps.CodTipoExp
        !Descripcion = mProps.DescripcionTipoExp
        !BandValida = mProps.bandValidaTipoExp
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.IdTipoEXp = !Idtipoexportacion
        .Close
    End With
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_ANCOMP", "AnexoComp Nuevo: " & _
                                                         mProps.CodComprobante, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_ANCOMP", "AnexoComp Modificado: " & _
                                                         mProps.CodComprobante, "DT"
    End If
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_ANEXOTIPOEXP
End Sub
Public Property Let CodTipoExp(ByVal value As String)
    ValidaCodigo value      'Valida si contiene caracteres no validos
    mProps.CodTipoExp = value
    mbooModificado = True
End Property
Public Property Get CodTipoExp() As String
    CodTipoExp = Trim$(mProps.CodTipoExp)
End Property
Public Property Let DescripcionTipoExp(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Anexos_DescripcionTipoExp", MSGERR_INVALIDO
    End If
    mProps.DescripcionTipoExp = value
    mbooModificado = True
End Property
Public Property Get DescripcionTipoExp() As String
    DescripcionTipoExp = Trim$(mProps.DescripcionTipoExp)
End Property
Public Property Let bandValidaTipoExp(ByVal value As Boolean)
    mProps.bandValidaTipoExp = value
    mbooModificado = True
End Property
Public Property Get bandValidaTipoExp() As Boolean
    bandValidaTipoExp = mProps.bandValidaTipoExp
End Property
Public Sub EliminarTipoExp()
    Dim sql As String
    Dim rs As Recordset
    'Verifica si ya existe un item usando el mismo tIPO dOCUMENTO
'    sql = "SELECT CodProvCli FROM PCProvCli " & _
'          " WHERE IDTipoExp = " & mProps.IdTipoEXp
'    Set rs = mProps.objEmpresa.OpenRecordset(sql)
'    If Not rs.EOF Then
'        Err.Raise ERR_NOELIMINABLE, "Anexo_TipoExp", MSGERR_NOELIMINABLE
'    End If
'    rs.Close
    'Elimina el registro
    sql = "DELETE FROM Anexo_TipoExportacion WHERE CodTipoExportacion= '" & mProps.CodTipoExp & "'"
    mProps.objEmpresa.Execute sql, True
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_TIDOC", "AnexoTipoExp Eliminado: " & _
                                                         mProps.CodTipoExp, "DT"
    Set rs = Nothing
    mbooEsNuevo = True
    mbooModificado = True
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_ANEXOTIPOEXP
End Sub


Public Property Let BandAutoImpresor(ByVal value As Boolean)
    mProps.BandAutoImpresor = value
    mbooModificado = True
End Property

Public Property Get BandAutoImpresor() As Boolean
    BandAutoImpresor = mProps.BandAutoImpresor
End Property

Public Property Let BandAutorizaActiva(ByVal value As Boolean)
    mProps.BandAutorizaActiva = value
    mbooModificado = True
End Property

Public Property Get BandAutorizaActiva() As Boolean
    BandAutorizaActiva = mProps.BandAutorizaActiva
End Property

'--------------------------------------------------------------------
Public Sub GrabarFormaPago()
    Dim sql As String, rs As Recordset
    
        'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
    sql = "SELECT * FROM Anexo_FormaPago" & _
          " WHERE  CodFormaPago  ='" & mProps.CodFormaPago & "'" & _
          "And Id =  " & mProps.IdFormaPago
          
   
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM  Anexo_FormaPago  WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        'ojo para poder cambiar un codigo se debe crear un id en c/tabla
        sql = "SELECT * FROM Anexo_FormaPago WHERE id =" & mProps.IdFormaPago
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
End If
    'Graba el registro en Pos
    With rs
        !CodFormaPago = mProps.CodFormaPago
        !Descripcion = mProps.DescripcionFormaPago
        !BandValida = mProps.bandValidaFormaPago
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.IdFormaPago = !id
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_ANFP", "AnexoComp Nuevo: " & _
                                                         mProps.CodComprobante, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_ANFP", "AnexoComp Modificado: " & _
                                                         mProps.CodComprobante, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_ANEXOFORMAPAGO
End Sub

Friend Sub RecuperarAnexoFormaPago(key As Variant)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT * FROM  Anexo_FormaPago WHERE CodFormaPago ='" & key & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        mProps.IdFormaPago = !id
        mProps.CodFormaPago = !CodFormaPago
        If Not IsNull(!Descripcion) Then mProps.DescripcionFormaPago = !Descripcion
        If Not IsNull(!BandValida) Then mProps.bandValidaFormaPago = !BandValida
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub EliminarFormaPago()
    Dim sql As String
    Dim rs As Recordset
    

    'Elimina el registro
    sql = "DELETE FROM Anexo_FormaPago WHERE CodFormaPago= '" & mProps.CodFormaPago & "'"
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_ANFP", "AnexoFormaPago Eliminado: " & _
                                                         mProps.CodFormaPago, "DT"
                                                         
    Set rs = Nothing
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_ANEXOFORMAPAGO
End Sub

Public Property Let CodFormaPago(ByVal value As String)
    ValidaCodigo value      'Valida si contiene caracteres no validos
    mProps.CodFormaPago = value
    mbooModificado = True
End Property

Public Property Get CodFormaPago() As String
    CodFormaPago = Trim$(mProps.CodFormaPago)
End Property

Public Property Let DescripcionFormaPago(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Anexos_DescripcionFormaPago", MSGERR_INVALIDO
    End If
    mProps.DescripcionFormaPago = value
    mbooModificado = True
End Property

Public Property Get DescripcionFormaPago() As String
    DescripcionFormaPago = Trim$(mProps.DescripcionFormaPago)
End Property

Public Property Let bandValidaFormaPago(ByVal value As Boolean)
    mProps.bandValidaFormaPago = value
    mbooModificado = True
End Property

Public Property Get bandValidaFormaPago() As Boolean
    bandValidaFormaPago = mProps.bandValidaFormaPago
End Property

'--------------------------------------------------------------------
Public Sub GrabarPais()
    Dim sql As String, rs As Recordset
    
        'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
    sql = "SELECT * FROM Anexo_Pais" & _
          " WHERE  CodPais  ='" & mProps.CodPais & "'" & _
          "And Id =  " & mProps.IdPais
          
   
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM  Anexo_Pais  WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        'ojo para poder cambiar un codigo se debe crear un id en c/tabla
        sql = "SELECT * FROM Anexo_Pais WHERE id =" & mProps.IdPais
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
End If
    'Graba el registro en Pos
    With rs
        !CodPais = mProps.CodPais
        !Descripcion = mProps.DescripcionPais
        !BandValida = mProps.bandValidaPais
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.IdPais = !id
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_ANFP", "AnexoComp Nuevo: " & _
                                                         mProps.CodComprobante, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_ANFP", "AnexoComp Modificado: " & _
                                                         mProps.CodComprobante, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_ANEXOPAIS
End Sub

Friend Sub RecuperarAnexoPais(key As Variant)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT * FROM  Anexo_Pais WHERE CodPais ='" & key & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        mProps.IdPais = !id
        mProps.CodPais = !CodPais
        If Not IsNull(!Descripcion) Then mProps.DescripcionPais = !Descripcion
        If Not IsNull(!BandValida) Then mProps.bandValidaPais = !BandValida
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub EliminarPais()
    Dim sql As String
    Dim rs As Recordset
    

    'Elimina el registro
    sql = "DELETE FROM Anexo_Pais WHERE CodPais= '" & mProps.CodPais & "'"
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_ANFP", "AnexoPais Eliminado: " & _
                                                         mProps.CodPais, "DT"
                                                         
    Set rs = Nothing
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_ANEXOPAIS
End Sub

Public Property Let CodPais(ByVal value As String)
    ValidaCodigo value      'Valida si contiene caracteres no validos
    mProps.CodPais = value
    mbooModificado = True
End Property

Public Property Get CodPais() As String
    CodPais = Trim$(mProps.CodPais)
End Property

Public Property Let DescripcionPais(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Anexos_DescripcionPais", MSGERR_INVALIDO
    End If
    mProps.DescripcionPais = value
    mbooModificado = True
End Property

Public Property Get DescripcionPais() As String
    DescripcionPais = Trim$(mProps.DescripcionPais)
End Property

Public Property Let bandValidaPais(ByVal value As Boolean)
    mProps.bandValidaPais = value
    mbooModificado = True
End Property

Public Property Get bandValidaPais() As Boolean
    bandValidaPais = mProps.bandValidaPais
End Property



Public Sub GrabarRetIVA()
    Dim sql As String, rs As Recordset
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
    sql = "SELECT * FROM Anexos_Retencion_IVA" & _
          " WHERE  codretencion  ='" & mProps.CodRetencion & "'" & _
          " AND Id = " & mProps.IdRetencion
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM  Anexos_Retencion_IVA WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM Anexos_Retencion_IVA WHERE id ='" & mProps.IdRetencion & "'"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
    End If
    'Graba el registro en Pos
    With rs
        !CodRetencion = mProps.CodRetencion
        !Descripcion = mProps.DescripcionRetIR
        !Porcentaje = mProps.Porcentaje
        !BandValida = mProps.bandValidaRet
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
         mProps.IdRetencion = !id
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_ANRTIVA", "AnexoRETIVA Nuevo: " & _
                                                         mProps.CodRetencion, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_ANRTIVA", "AnexoRETIVA Modificado: " & _
                                                         mProps.CodRetencion, "DT"
    End If
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_ANEXORETIVA
End Sub


Friend Sub RecuperarAnexoRetIVA(key As Variant)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT * FROM  Anexos_retencion_iva WHERE  codretencion ='" & key & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        mProps.IdRetencion = !id
        mProps.CodRetencion = !CodRetencion
        If Not IsNull(!Descripcion) Then mProps.DescripcionRetIR = !Descripcion
        If Not IsNull(!Porcentaje) Then mProps.Porcentaje = !Porcentaje
        If Not IsNull(!BandValida) Then mProps.bandValidaRet = !BandValida
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub EliminarRetIVA()
    Dim sql As String
    Dim rs As Recordset

    sql = "SELECT CodRetencion FROM Anexos_Retencion_Iva" & _
          " WHERE codRetencion= '" & mProps.CodRetencion & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'Elimina el registro
    sql = "DELETE FROM Anexos_Retencion_IVA WHERE codRetencion= '" & mProps.CodRetencion & "'"
    mProps.objEmpresa.Execute sql, True
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_ANRETIVA", "AnexoRETIVA Eliminado: " & _
                                                         mProps.CodRetencion, "DT"
    Set rs = Nothing
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_ANEXORETIVA
End Sub

'---------------------- anexo - ice

'--------------------------------------------------------------------
Public Sub GrabarICE()
    Dim sql As String, rs As Recordset
    
        'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
    sql = "SELECT * FROM Anexo_ICE" & _
          " WHERE  CodICE  ='" & mProps.CodICE & "'" & _
          "And Id =  " & mProps.IdICE
          
   
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM  Anexo_ICE  WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        'ojo para poder cambiar un codigo se debe crear un id en c/tabla
        sql = "SELECT * FROM Anexo_ICE WHERE id =" & mProps.IdICE
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
End If
    'Graba el registro en Pos
    With rs
        !CodICE = mProps.CodICE
        !Descripcion = mProps.DescripcionICE
        !BandValida = mProps.bandValidaICE
        !CodClasifica = mProps.CodClasifica
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.IdICE = !id
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_ANE", "AnexoComp Nuevo: " & _
                                                         mProps.CodComprobante, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_ANE", "AnexoComp Modificado: " & _
                                                         mProps.CodComprobante, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_ANEXOICE
End Sub

Friend Sub RecuperarAnexoICE(key As Variant)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT * FROM  Anexo_ICE WHERE CodICE ='" & key & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        mProps.IdICE = !id
        mProps.CodICE = !CodICE
        If Not IsNull(!Descripcion) Then mProps.DescripcionICE = !Descripcion
        If Not IsNull(!BandValida) Then mProps.bandValidaICE = !BandValida
        If Not IsNull(!CodClasifica) Then mProps.CodClasifica = !CodClasifica
        
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub EliminarICE()
    Dim sql As String
    Dim rs As Recordset
    

    'Elimina el registro
    sql = "DELETE FROM Anexo_ICE WHERE CodICE= '" & mProps.CodICE & "'"
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_ANFP", "AnexoICE Eliminado: " & _
                                                         mProps.CodICE, "DT"
                                                         
    Set rs = Nothing
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_ANEXOICE
End Sub

Public Property Let CodICE(ByVal value As String)
    ValidaCodigo value      'Valida si contiene caracteres no validos
    mProps.CodICE = value
    mbooModificado = True
End Property

Public Property Get CodICE() As String
    CodICE = Trim$(mProps.CodICE)
End Property

Public Property Let DescripcionICE(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "Anexos_DescripcionICE", MSGERR_INVALIDO
    End If
    mProps.DescripcionICE = value
    mbooModificado = True
End Property

Public Property Get DescripcionICE() As String
    DescripcionICE = Trim$(mProps.DescripcionICE)
End Property

Public Property Let bandValidaICE(ByVal value As Boolean)
    mProps.bandValidaICE = value
    mbooModificado = True
End Property

Public Property Get bandValidaICE() As Boolean
    bandValidaICE = mProps.bandValidaICE
End Property



Public Property Let CodClasifica(ByVal value As String)
    If Len(value) > 2 Then
        Err.Raise ERR_INVALIDO, "Anexos_CodClasifica", MSGERR_INVALIDO
    End If

    mProps.CodClasifica = value
    mbooModificado = True
End Property

Public Property Get CodClasifica() As String
    CodClasifica = Trim$(mProps.CodClasifica)
End Property

