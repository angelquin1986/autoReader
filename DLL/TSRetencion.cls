VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TSRetencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Empresa"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Private Type T_PROPS
    IdRetencion As Long
    CodRetencion As String
    Descripcion As String
    CodSRI As String 'Agregado AUC 01/09/2005
    Porcentaje As Single
    IdCuentaActivo As Long
    CodCuentaActivo As String
    IdCuentaPasivo As Long
    CodCuentaPasivo As String
    BandValida As Boolean
    'BandIR As Boolean 'AUC 30/03/07
    BandIVA As Boolean
    IdAnexo As Long    'AUC 02/04/07
    CodAnexo As String 'AUC 02/04/07
    CodF104 As String 'Agregado jeaa 23/04/2007
    BandBienes As Boolean 'Agregado jeaa 25/09/2007
    BandServicios As Boolean 'Agregado jeaa 25/09/2007
    BandCompras As Boolean 'Agregado jeaa 03/07/2008
    BandVentas As Boolean 'Agregado jeaa 03/07/2008
    fechagrabado As Date
    
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
    sql = "SELECT CodRetencion FROM TSRetencion " & _
          "WHERE (CodRetencion='" & mProps.CodRetencion & "') AND " & _
                "(IdRetencion<>" & mProps.IdRetencion & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "TSRetencion", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM TSRetencion WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM TSRetencion WHERE IdRetencion=" & mProps.IdRetencion
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "TSRetencion.Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en Pos
    With rs

        If Not !CodRetencion = mProps.CodRetencion Then
            mProps.objEmpresa.GrabaAccionModulo !CodRetencion, mProps.CodRetencion, "CodRetencion", "MOD_RET", "DT"
        End If
        !CodRetencion = mProps.CodRetencion
         
        If Not !Descripcion = mProps.Descripcion Then
            mProps.objEmpresa.GrabaAccionModulo !Descripcion, mProps.Descripcion, "Descripcion Retención", "MOD_RET", "DT"
        End If
         !Descripcion = mProps.Descripcion
         !CodSRI = mProps.CodSRI
        'jeaa 07/01/2007
        If Not !Porcentaje = mProps.Porcentaje Then
            mProps.objEmpresa.GrabaAccionModulo Round(!Porcentaje * 100, 4), mProps.Porcentaje * 100, "Porcentaje", "MOD_RET", "DT"
        End If
         !Porcentaje = mProps.Porcentaje
         !IdCuentaActivo = mProps.IdCuentaActivo
         !IdCuentaPasivo = mProps.IdCuentaPasivo
         !BandValida = mProps.BandValida
        If Not !BandIVA = mProps.BandIVA Then
            mProps.objEmpresa.GrabaAccionModulo !BandIVA, mProps.BandIVA, "BandIVA", "MOD_RET", "DT"
        End If
         !BandIVA = mProps.BandIVA 'AUC 30/03/07
         !IdAnexo = mProps.IdAnexo 'AUC 02/04/07
         !CodF104 = mProps.CodF104 'JEAA 23/04/2007
         !BandBienes = mProps.BandBienes 'jeaa 25/09/2007
         !BandServicios = mProps.BandServicios 'jeaa 25/09/2007
         !BandCompras = mProps.BandCompras 'jeaa 04/07/2008
         !BandVentas = mProps.BandVentas 'jeaa 04/07/2008
         !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IdRetencion = !IdRetencion
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_RET", "Retención Nueva: " & _
                                                         mProps.CodRetencion, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_RET", "Retención Modificada: " & _
                                                         mProps.CodRetencion, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_RETENCION
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT tsr.*, " & _
                "cta.CodCuenta AS CodCuentaActivo, " & _
                "ctp.CodCuenta AS CodCuentaPasivo,CASE WHEN bandIVA= 1 THEN aneiva.codretencion ELSE aneir.codretencion END as codAnexo  " & _
          "FROM CTCuenta ctp RIGHT JOIN " & _
                    "(CTCuenta cta RIGHT JOIN TSRetencion tsr " & _
                    "ON cta.IdCuenta = tsr.IdCuentaActivo) " & _
                "ON ctp.IdCuenta = tsr.IdCuentaPasivo " & _
                "LEFT JOIN anexos_retencion_ir aneir on  aneir.id = tsr.idanexo " & _
                "LEFT JOIN anexos_retencion_iva aneiva on  aneiva.id = tsr.idanexo "
    If VarType(key) = vbString Then
        sql = sql & "WHERE tsr.CodRetencion='" & key & "'"
    Else
        sql = "WHERE tsr.IdRetencion=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos
    With rs
        mProps.IdRetencion = !IdRetencion
        mProps.CodRetencion = !CodRetencion
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!Porcentaje) Then mProps.Porcentaje = !Porcentaje
        If Not IsNull(!IdCuentaActivo) Then mProps.IdCuentaActivo = !IdCuentaActivo
        If Not IsNull(!CodCuentaActivo) Then mProps.CodCuentaActivo = !CodCuentaActivo
        If Not IsNull(!IdCuentaPasivo) Then mProps.IdCuentaPasivo = !IdCuentaPasivo
        If Not IsNull(!CodCuentaPasivo) Then mProps.CodCuentaPasivo = !CodCuentaPasivo
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!CodSRI) Then mProps.CodSRI = !CodSRI 'Agregado AUC  01/09/2005
'        If Not IsNull(!BandIR) Then mProps.BandIR = !BandIR  'Agregado AUC  30/03/2007
        If Not IsNull(!BandIVA) Then mProps.BandIVA = !BandIVA  'Agregado AUC  30/03/2007
        If Not IsNull(!CodAnexo) Then mProps.CodAnexo = !CodAnexo
        If Not IsNull(!IdAnexo) Then mProps.IdAnexo = !IdAnexo 'AUC 02/04/07
        If Not IsNull(!CodF104) Then mProps.CodF104 = !CodF104 'Agregado JEAA 23/04/2007
        If Not IsNull(!BandBienes) Then mProps.BandBienes = !BandBienes 'Agregado JEAA 25/09/2007
        If Not IsNull(!BandServicios) Then mProps.BandServicios = !BandServicios 'Agregado JEAA 25/09/2007
        If Not IsNull(!BandCompras) Then mProps.BandCompras = !BandCompras 'Agregado JEAA 04/07/2008
        If Not IsNull(!BandVentas) Then mProps.BandVentas = !BandVentas 'Agregado JEAA 04/07/2008
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset

    'Verifica si existe Kardex del retencion
    sql = "SELECT IdRetencion FROM TSKardexRet WHERE IdRetencion=" & mProps.IdRetencion
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "TSRetencion.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing

    'Elimina la retencion
    sql = "DELETE FROM TSRetencion WHERE IdRetencion=" & mProps.IdRetencion
    mProps.objEmpresa.Execute sql, True

    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_RET", "Retención Eliminada: " & _
                                                         mProps.CodRetencion, "DT"

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_RETENCION
End Sub

Public Function Clone() As TSRetencion
    Dim nuevo As TSRetencion
    
    Set nuevo = New TSRetencion
    'Copia las propiedades al nuevo objeto
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdRetencion = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property



Public Property Let CodRetencion(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "TSRetencion.CodRetencion", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodRetencion = value
    mbooModificado = True
End Property

Public Property Get CodRetencion() As String
    CodRetencion = Trim$(mProps.CodRetencion)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 40 Then
        Err.Raise ERR_INVALIDO, "TSRetencion.Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let IdCuentaActivo(ByVal value As Long)
    mProps.IdCuentaActivo = value
    mbooModificado = True
End Property

Public Property Get IdCuentaActivo() As Long
    IdCuentaActivo = mProps.IdCuentaActivo
End Property


Public Property Let CodCuentaActivo(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCuentaActivo Then
        'Actualiza IdCuentaActivo también
        sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "TSRetencion.CodCuentaActivo", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            IdCuentaActivo = rs!IdCuenta
        End If
        rs.Close
    End If

    Set rs = Nothing
    mProps.CodCuentaActivo = value
    mbooModificado = True
End Property

Public Property Get CodCuentaActivo() As String
    CodCuentaActivo = Trim$(mProps.CodCuentaActivo)
End Property


Public Property Let IdCuentaPasivo(ByVal value As Long)
    mProps.IdCuentaPasivo = value
    mbooModificado = True
End Property

Public Property Get IdCuentaPasivo() As Long
    IdCuentaPasivo = mProps.IdCuentaPasivo
End Property


Public Property Let CodCuentaPasivo(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCuentaPasivo Then
        'Actualiza IdCuentaPasivo también
        sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "TSRetencion.CodCuentaPasivo", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            IdCuentaPasivo = rs!IdCuenta
        End If
        rs.Close
    End If

    Set rs = Nothing
    mProps.CodCuentaPasivo = value
    mbooModificado = True
End Property

Public Property Get CodCuentaPasivo() As String
    CodCuentaPasivo = Trim$(mProps.CodCuentaPasivo)
End Property




Public Property Let Porcentaje(ByVal value As Single)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "TSRetencion.Porcentaje", MSGERR_INVALIDO
    End If
    
    mProps.Porcentaje = value
    mbooModificado = True
End Property

Public Property Get Porcentaje() As Single
    Porcentaje = mProps.Porcentaje
End Property

Public Property Let BandValida(ByVal value As Boolean)
    mProps.BandValida = value
    mbooModificado = True
End Property

Public Property Get BandValida() As Boolean
    BandValida = mProps.BandValida
End Property




Public Property Get fechagrabado() As Date
    fechagrabado = mProps.fechagrabado
End Property

Public Property Get EsNuevo() As Boolean
    EsNuevo = mbooEsNuevo
End Property

Public Property Get Modificado() As Boolean
    Modificado = mbooModificado
End Property


Private Sub Class_Initialize()
    With mProps
        .BandValida = True
    End With
    
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub
Public Property Let CodSRI(ByVal value As String) 'Agregado AUC 01/09/2005
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "TSRetencion.CodSRI", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodSRI = value
    mbooModificado = True
End Property

Public Property Get CodSRI() As String 'Agregado AUC 01/09/2005
    CodSRI = Trim$(mProps.CodSRI)
End Property
'
''AUC 30/03/07
'Public Property Let BandIR(ByVal value As Boolean)
'    mProps.BandIR = value
'     mbooModificado = True
'End Property
'
'Public Property Get BandIR() As Boolean
'    BandIR = mProps.BandIR
'End Property
Public Property Let BandIVA(ByVal value As Boolean)
    mProps.BandIVA = value
     mbooModificado = True
End Property

Public Property Get BandIVA() As Boolean
    BandIVA = mProps.BandIVA
End Property
'AUC 02/04/07
Public Property Let IdAnexo(ByVal value As Long)
    mProps.IdAnexo = value
    mbooModificado = True
End Property
Public Property Get IdAnexo() As Long
    IdAnexo = mProps.IdAnexo
End Property

Public Property Let CodAnexo(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodAnexo Then
        'Actualiza IdCuentaActivo también
        If Me.BandIVA Then
            sql = "SELECT Id FROM Anexos_Retencion_IVA WHERE Codretencion='" & value & "'"
        Else
            sql = "SELECT Id FROM Anexos_Retencion_IR WHERE Codretencion='" & value & "'"
        End If
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "ANEXO_CODRETENCION", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            IdAnexo = rs!id
        End If
        rs.Close
    End If

    Set rs = Nothing
    mProps.CodAnexo = value
    mbooModificado = True
End Property

Public Property Get CodAnexo() As String
    CodAnexo = Trim$(mProps.CodAnexo)
End Property

Public Property Let CodF104(ByVal value As String) 'Agregado JEAA 23/04/2007
    If Len(value) > 3 Then
        Err.Raise ERR_INVALIDO, "TSRetencion.CodF104", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    mProps.CodF104 = value
    mbooModificado = True
End Property

Public Property Get CodF104() As String 'Agregado JEAA 23/04/2007
    CodF104 = Trim$(mProps.CodF104)
End Property

Public Property Let BandBienes(ByVal value As Boolean)
    mProps.BandBienes = value
     mbooModificado = True
End Property

Public Property Get BandBienes() As Boolean
    BandBienes = mProps.BandBienes
End Property

Public Property Let BandServicios(ByVal value As Boolean)
    mProps.BandServicios = value
     mbooModificado = True
End Property

Public Property Get BandServicios() As Boolean
    BandServicios = mProps.BandServicios
End Property

Public Property Let BandCompras(ByVal value As Boolean)
    mProps.BandCompras = value
     mbooModificado = True
End Property

Public Property Get BandCompras() As Boolean
    BandCompras = mProps.BandCompras
End Property

Public Property Let BandVentas(ByVal value As Boolean)
    mProps.BandVentas = value
     mbooModificado = True
End Property

Public Property Get BandVentas() As Boolean
    BandVentas = mProps.BandVentas
End Property


Friend Sub RecuperarxSRI(key As Variant)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT tsr.*, " & _
                "cta.CodCuenta AS CodCuentaActivo, " & _
                "ctp.CodCuenta AS CodCuentaPasivo,CASE WHEN bandIVA= 1 THEN aneiva.codretencion ELSE aneir.codretencion END as codAnexo  " & _
          "FROM CTCuenta ctp RIGHT JOIN " & _
                    "(CTCuenta cta RIGHT JOIN TSRetencion tsr " & _
                    "ON cta.IdCuenta = tsr.IdCuentaActivo) " & _
                "ON ctp.IdCuenta = tsr.IdCuentaPasivo " & _
                "LEFT JOIN anexos_retencion_ir aneir on  aneir.id = tsr.idanexo " & _
                "LEFT JOIN anexos_retencion_iva aneiva on  aneiva.id = tsr.idanexo "
    If VarType(key) = vbString Then
        sql = sql & " WHERE aneir.CodRetencion='" & key & "' or aneiva.CodRetencion='" & key & "'"
    Else
        sql = "WHERE tsr.IdRetencion=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos
    With rs
        mProps.IdRetencion = !IdRetencion
        mProps.CodRetencion = !CodRetencion
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!Porcentaje) Then mProps.Porcentaje = !Porcentaje
        If Not IsNull(!IdCuentaActivo) Then mProps.IdCuentaActivo = !IdCuentaActivo
        If Not IsNull(!CodCuentaActivo) Then mProps.CodCuentaActivo = !CodCuentaActivo
        If Not IsNull(!IdCuentaPasivo) Then mProps.IdCuentaPasivo = !IdCuentaPasivo
        If Not IsNull(!CodCuentaPasivo) Then mProps.CodCuentaPasivo = !CodCuentaPasivo
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!CodSRI) Then mProps.CodSRI = !CodSRI 'Agregado AUC  01/09/2005
'        If Not IsNull(!BandIR) Then mProps.BandIR = !BandIR  'Agregado AUC  30/03/2007
        If Not IsNull(!BandIVA) Then mProps.BandIVA = !BandIVA  'Agregado AUC  30/03/2007
        If Not IsNull(!CodAnexo) Then mProps.CodAnexo = !CodAnexo
        If Not IsNull(!IdAnexo) Then mProps.IdAnexo = !IdAnexo 'AUC 02/04/07
        If Not IsNull(!CodF104) Then mProps.CodF104 = !CodF104 'Agregado JEAA 23/04/2007
        If Not IsNull(!BandBienes) Then mProps.BandBienes = !BandBienes 'Agregado JEAA 25/09/2007
        If Not IsNull(!BandServicios) Then mProps.BandServicios = !BandServicios 'Agregado JEAA 25/09/2007
        If Not IsNull(!BandCompras) Then mProps.BandCompras = !BandCompras 'Agregado JEAA 04/07/2008
        If Not IsNull(!BandVentas) Then mProps.BandVentas = !BandVentas 'Agregado JEAA 04/07/2008
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Friend Sub RecuperarxSRIxIVA(key As Variant, BandIVA As Boolean)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT tsr.*, " & _
                "cta.CodCuenta AS CodCuentaActivo, " & _
                "ctp.CodCuenta AS CodCuentaPasivo,"
    If BandIVA Then
        sql = sql & "  aneiva.codretencion  as codAnexo  "
    Else
        sql = sql & " aneir.codretencion  as codAnexo  "
    End If
    sql = sql & " FROM CTCuenta ctp RIGHT JOIN " & _
                    "(CTCuenta cta RIGHT JOIN TSRetencion tsr " & _
                    "ON cta.IdCuenta = tsr.IdCuentaActivo) " & _
                "ON ctp.IdCuenta = tsr.IdCuentaPasivo "
    If BandIVA Then
        sql = sql & " LEFT JOIN anexos_retencion_iva aneiva on  aneiva.id = tsr.idanexo "
    Else
        sql = sql & " LEFT JOIN anexos_retencion_ir aneir on  aneir.id = tsr.idanexo "
    End If
                
    If VarType(key) = vbString Then
        If BandIVA Then
            sql = sql & " WHERE bandventas=1 and bandiva=1 and aneiva.CodRetencion='" & key & "'"
        Else
            sql = sql & " WHERE bandventas=1 and bandiva=0 and aneir.CodRetencion='" & key & "' "
        End If
    Else
        sql = "WHERE tsr.IdRetencion=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos
    With rs
        mProps.IdRetencion = !IdRetencion
        mProps.CodRetencion = !CodRetencion
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!Porcentaje) Then mProps.Porcentaje = !Porcentaje
        If Not IsNull(!IdCuentaActivo) Then mProps.IdCuentaActivo = !IdCuentaActivo
        If Not IsNull(!CodCuentaActivo) Then mProps.CodCuentaActivo = !CodCuentaActivo
        If Not IsNull(!IdCuentaPasivo) Then mProps.IdCuentaPasivo = !IdCuentaPasivo
        If Not IsNull(!CodCuentaPasivo) Then mProps.CodCuentaPasivo = !CodCuentaPasivo
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!CodSRI) Then mProps.CodSRI = !CodSRI 'Agregado AUC  01/09/2005
'        If Not IsNull(!BandIR) Then mProps.BandIR = !BandIR  'Agregado AUC  30/03/2007
        If Not IsNull(!BandIVA) Then mProps.BandIVA = !BandIVA  'Agregado AUC  30/03/2007
        If Not IsNull(!CodAnexo) Then mProps.CodAnexo = !CodAnexo
        If Not IsNull(!IdAnexo) Then mProps.IdAnexo = !IdAnexo 'AUC 02/04/07
        If Not IsNull(!CodF104) Then mProps.CodF104 = !CodF104 'Agregado JEAA 23/04/2007
        If Not IsNull(!BandBienes) Then mProps.BandBienes = !BandBienes 'Agregado JEAA 25/09/2007
        If Not IsNull(!BandServicios) Then mProps.BandServicios = !BandServicios 'Agregado JEAA 25/09/2007
        If Not IsNull(!BandCompras) Then mProps.BandCompras = !BandCompras 'Agregado JEAA 04/07/2008
        If Not IsNull(!BandVentas) Then mProps.BandVentas = !BandVentas 'Agregado JEAA 04/07/2008
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub


