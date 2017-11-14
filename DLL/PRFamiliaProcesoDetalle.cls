VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PRFamiliaProcesoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdProceso As Long
    CodProceso As String
    Descripcion As String
    Unidad As String
    Cantidad As Currency
    CostoUltimoIngreso As Currency 'AUC 22/12/05
    TotalReceta As Currency 'AUC 22/12/05
    pVenta As Currency 'AUC 23/03/06
    xCuanto As Currency 'AUC 26/09/07 para promociones
    Precio As Currency 'AUC 28/09/07 para promociones
    objPRProceso As PRProceso          'Referencia al objeto padre
    objGNtrans As GNTrans
    
End Type
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Private mProps As T_PROPS

Friend Property Set PRProceso(obj As PRProceso)
    Set mProps.objPRProceso = obj
End Property

Public Property Get PRProceso() As PRProceso
    Set PRProceso = mProps.objPRProceso
End Property

Friend Property Let IdProceso(ByVal value As Long)
    mProps.IdProceso = value
End Property

Friend Property Get IdProceso() As Long
    IdProceso = mProps.IdProceso
End Property


Public Property Let CodProceso(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PRProceso_CodProceso", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodProceso = value
    mProps.objPRProceso.Modificado = True
    'mbooModificado = True
End Property

Public Property Get CodProceso() As String
    CodProceso = Trim$(mProps.CodProceso)
End Property


Public Property Let Cantidad(ByVal value As Currency)
    mProps.Cantidad = value
    mProps.objPRProceso.Modificado = True   'trabaja con  la bandera modificado del papa
End Property

Public Property Get Cantidad() As Currency
    Cantidad = mProps.Cantidad
End Property

Private Sub Class_Terminate()
    Set mProps.objPRProceso = Nothing
End Sub


Public Property Get Descripcion() As String 'Propiedad  solo de lectura
    
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PRProceso_Descripcion", MSGERR_INVALIDO
    End If
    
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, ";") > 0 Then
        Err.Raise ERR_INVALIDO, "PRProceso_Descripcion", _
            "El caracter ';' no es válido para la descripción."
    End If
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, "|") > 0 Then
        Err.Raise ERR_INVALIDO, "PRProceso_Descripcion", _
            "El caracter '|' no es válido para la descripción."
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property



Public Property Get Unidad() As String  'Propiedad  solo de lectura
    Unidad = Trim$(mProps.Unidad)
End Property

Public Property Let Unidad(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "PRProceso_Unidad", MSGERR_INVALIDO
    End If
    
    mProps.Unidad = value
    mbooModificado = True
End Property

 'key puede ser IdProceso  o CodProceso
Public Sub Recuperar(key As Variant)
   Dim sql As String, i As Integer
    Dim rs As Recordset
    Dim cond As String
    cond = " IV.CodProceso='" & key & "'"
'    If PRProceso.Empresa.GNOpcion.ObtenerValor("itemPreparacion") = 0 Then 'AUC toma el p.costo
        If VarType(key) = vbString Then
            sql = "Select IV.IdProceso,  IV.CodProceso, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
            "FROM PRProceso IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad WHERE" & cond
        Else
            sql = "Select IV.IdProceso,  IV.CodProceso, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
            "FROM PRProceso IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad  WHERE " & cond
        End If
'    ElseIf PRProceso.Empresa.GNOpcion.ObtenerValor("itemPreparacion") = 1 Then 'AUC toma el p.venta
'        If VarType(key) = vbString Then
'            sql = "Select IV.IdProceso,  IV.CodProceso, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
'            "FROM PRProceso IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad WHERE" & cond
'        Else
'            sql = "Select IV.IdProceso,  IV.CodProceso, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
'            "FROM PRProceso IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad  WHERE " & cond
'        End If
'    End If
    Set rs = Me.PRProceso.Empresa.OpenRecordset(sql)
    If rs.EOF Then 'AUC cambiado antes estuvo NOT (RS.EOF)
       Err.Raise ERR_NOHAYCODIGO, "IvFamiliaDetalle", MSGERR_NOHAYCODIGO
    Else
        mProps.IdProceso = rs!IdProceso
        mProps.CodProceso = rs!CodProceso
        mProps.Descripcion = rs!Descripcion
        If Not IsNull(rs!CodUnidad) Then mProps.Unidad = rs!CodUnidad
'        If PRProceso.Empresa.GNOpcion.ObtenerValor("itemPreparacion") = 0 Then 'AUC toma el p.costo
            If Not IsNull(rs!CostoUltimoIngreso) Then mProps.CostoUltimoIngreso = rs!CostoUltimoIngreso 'AUC 22/12/05
'        ElseIf PRProceso.Empresa.GNOpcion.ObtenerValor("itemPreparacion") = 1 Then 'AUC toma el p.venta
'            mProps.pVenta = ObtienePrecio(mProps.CodProceso)
'        End If
  End If
End Sub
'AUC 22/12/05
Public Property Let CostoUltimoIngreso(ByVal value As Currency)
    mProps.CostoUltimoIngreso = value
    mProps.objPRProceso.Modificado = True   'trabaja con  la bandera modificado del papa
End Property
'AUC 22/12/05
Public Property Get CostoUltimoIngreso() As Currency
    CostoUltimoIngreso = mProps.CostoUltimoIngreso
End Property
'AUC 22/12/05
Public Property Let TotalReceta(ByVal value As Currency)
    mProps.TotalReceta = value
    mProps.objPRProceso.Modificado = True   'trabaja con  la bandera modificado del papa
End Property
'AUC 22/12/05
Public Property Get TotalReceta() As Currency
    TotalReceta = mProps.TotalReceta
End Property

Public Sub RecuperarPreparacion(key As Variant, Optional SiPreparacion As String)
    Dim sql As String, i As Integer
    Dim rs As Recordset
    'select IdProceso,  CodProceso, Descripcion, Unidad   from PRProceso
    If VarType(key) = vbString Then
        sql = "Select IV.IdProceso,  IV.CodProceso, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
        " FROM PRProceso IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad WHERE CodProceso='" & key & "'"
    Else
        sql = "Select IV.IdProceso,  IV.CodProceso, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
        " FROM PRProceso IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad  WHERE IdProceso='" & key & "'"
    End If
    Set rs = Me.PRProceso.Empresa.OpenRecordset(sql)
    If rs.EOF Then
        Err.Raise ERR_NOHAYCODIGO, "IvFamiliaDetalle", MSGERR_NOHAYCODIGO
    Else
        mProps.IdProceso = rs!IdProceso
        mProps.CodProceso = rs!CodProceso
        mProps.Descripcion = rs!Descripcion
        If Not IsNull(rs!CodUnidad) Then mProps.Unidad = rs!CodUnidad
        If PRProceso.Empresa.GNOpcion.ObtenerValor("itemPreparacion") = 0 Then 'AUC toma el p.costo
            If Not IsNull(rs!CostoUltimoIngreso) Then mProps.CostoUltimoIngreso = rs!CostoUltimoIngreso 'AUC 22/12/05
        ElseIf PRProceso.Empresa.GNOpcion.ObtenerValor("itemPreparacion") = 1 Then 'AUC toma el p.venta
            'mProps.pVenta = ObtienePrecio(mProps.CodProceso)
        End If
        'If Not IsNull(rs!CostoUltimoIngreso) Then mProps.CostoUltimoIngreso = rs!CostoUltimoIngreso 'AUC 22/12/05
    End If
End Sub
Friend Property Set GNTrans(obj As GNTrans)
    Set mProps.objGNtrans = obj
End Property
Public Property Get GNTrans() As GNTrans
    Set GNTrans = mProps.objGNtrans
End Property
'AUC 23/03/06
Public Property Let pVenta(ByVal value As Currency)
    mProps.pVenta = value
   mProps.objPRProceso.Modificado = True   'trabaja con  la bandera modificado del papa
End Property
'AUC 23/03/06
Public Property Get pVenta() As Currency
    pVenta = mProps.pVenta
End Property

'Public Function ObtienePrecio(cod As String) As Currency
'Dim item As PRProceso
'Dim ix As Integer
'Set item = PRProceso.Empresa.RecuperaPRProceso(cod)
'With item
'        Select Case PRProceso.Empresa.GNOpcion.ObtenerValor("Precio")
'            Case "P1"
'                ix = 1
'            Case "P2"
'                ix = 2
'            Case "P3"
'                ix = 3
'            Case "P4"
'                ix = 4
'        End Select
'     ObtienePrecio = .Precio(ix) + (.Precio(ix) * .PorcentajeIVA) '*  para sacar el precio MAS iva
'End With
'Set item = Nothing
'End Function
'key puede ser CodProceso o Descripcion
Public Sub RecuperarPreparacionDescripcion(key As Variant)
   Dim sql As String, i As Integer
    Dim rs As Recordset
    Dim cond As String
        cond = " IV.descripcion='" & key & "'"
'    If PRProceso.Empresa.GNOpcion.ObtenerValor("itemPreparacion") = 0 Then 'AUC toma el p.costo
       If VarType(key) = vbString Then
            sql = "Select IV.IdProceso,  IV.CodProceso, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
            "FROM PRProceso IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad WHERE" & cond
        Else
            sql = "Select IV.IdProceso,  IV.CodProceso, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
            "FROM PRProceso IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad  WHERE " & cond
        End If
'    ElseIf PRProceso.Empresa.GNOpcion.ObtenerValor("itemPreparacion") = 1 Then 'AUC toma el p.venta
'        If VarType(key) = vbString Then
'            sql = "Select IV.IdProceso,  IV.CodProceso, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
'            "FROM PRProceso IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad WHERE" & cond
'        Else
'            sql = "Select IV.IdProceso,  IV.CodProceso, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
'            "FROM PRProceso IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad  WHERE " & cond
'        End If
'    End If
   Set rs = Me.PRProceso.Empresa.OpenRecordset(sql)
    If rs.EOF Then 'AUC cambiado antes estuvo NOT (RS.EOF)
       Err.Raise ERR_NOHAYCODIGO, "IvFamiliaDetalle", MSGERR_NOHAYCODIGO
    Else
        mProps.IdProceso = rs!IdProceso
        mProps.CodProceso = rs!CodProceso
        mProps.Descripcion = rs!Descripcion
        If Not IsNull(rs!CodUnidad) Then mProps.Unidad = rs!CodUnidad
'        If PRProceso.Empresa.GNOpcion.ObtenerValor("itemPreparacion") = 0 Then 'AUC toma el p.costo
           If Not IsNull(rs!CostoUltimoIngreso) Then mProps.CostoUltimoIngreso = rs!CostoUltimoIngreso 'AUC 22/12/05
 '       ElseIf PRProceso.Empresa.GNOpcion.ObtenerValor("itemPreparacion") = 1 Then 'AUC toma el p.venta
  '          mProps.pVenta = ObtienePrecio(mProps.CodProceso)
   '     End If
  End If
End Sub
'AUC 26/09/07 para promociones
Public Property Let xCuanto(ByVal value As Currency)
    mProps.xCuanto = value
    mProps.objPRProceso.Modificado = True   'trabaja con  la bandera modificado del papa
End Property
Public Property Get xCuanto() As Currency
    xCuanto = mProps.xCuanto
End Property
'AUC 26/09/07
Public Sub RecuperarPromocion(key As Variant, Optional SiPreparacion As String)
    Dim sql As String, i As Integer
    Dim rs As Recordset
    'select IdProceso,  CodProceso, Descripcion, Unidad   from PRProceso
    If VarType(key) = vbString Then
        sql = "Select IV.IdProceso,  IV.CodProceso, IV.Descripcion, U.CodUnidad " & _
        " FROM PRProceso IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad WHERE CodProceso='" & key & "'"
    Else
        sql = "Select IV.IdProceso,  IV.CodProceso, IV.Descripcion, U.CodUnidad " & _
        " FROM PRProceso IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad  WHERE IdProceso='" & key & "'"
    End If
    Set rs = Me.PRProceso.Empresa.OpenRecordset(sql)
    If rs.EOF Then
        Err.Raise ERR_NOHAYCODIGO, "IvFamiliaDetalle", MSGERR_NOHAYCODIGO
    Else
        mProps.IdProceso = rs!IdProceso
        mProps.CodProceso = rs!CodProceso
        mProps.Descripcion = rs!Descripcion
        mProps.Cantidad = rs!Cantidad
        mProps.xCuanto = rs!xCuanto
        mProps.Precio = rs!Precio
    End If
End Sub
'AUC 26/09/07 para promociones
Public Property Let Precio(ByVal value As Currency)
    mProps.Precio = value
    mProps.objPRProceso.Modificado = True   'trabaja con  la bandera modificado del papa
End Property
Public Property Get Precio() As Currency
    Precio = mProps.Precio
End Property



