VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVFamiliaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"IVInventario"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type T_PROPS
    IdInventario As Long
    CodInventario As String
    Descripcion As String
    Unidad As String
    Cantidad As Currency
    CostoUltimoIngreso As Currency 'AUC 22/12/05
    TotalReceta As Currency 'AUC 22/12/05
    pVenta As Currency 'AUC 23/03/06
    BandPrincipal As Boolean 'AUC 09/05/06 campo restriccion para armar la formula
    BandModificar As Boolean 'AUC 25/05/06 Campo para modificar
    xCuanto As Currency 'AUC 26/09/07 para promociones
    TarifaJornal As Currency  'AUC 05/07/07 para rubros
    Rendimiento As Currency  'AUC 05/07/07 para rubros
    Orden As Integer
    Formula As String
    Precio As Currency 'AUC 28/09/07 para promociones
    
    objIVInventario As IVInventario          'Referencia al objeto padre
    Tipo As Integer
    CostoReferencial As Currency
    idproceso As Long
    CodProceso As String
    CodGrupo6 As String
    objGNtrans As GNTrans
    
End Type
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Private mProps As T_PROPS

Friend Property Set IVInventario(obj As IVInventario)
    Set mProps.objIVInventario = obj
End Property

Public Property Get IVInventario() As IVInventario
    Set IVInventario = mProps.objIVInventario
End Property

Friend Property Let IdInventario(ByVal value As Long)
    mProps.IdInventario = value
End Property

Friend Property Get IdInventario() As Long
    IdInventario = mProps.IdInventario
End Property


Public Property Let CodInventario(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "IVInventario_CodInventario", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodInventario = value
    mProps.objIVInventario.Modificado = True
    'mbooModificado = True
End Property

Public Property Get CodInventario() As String
    CodInventario = Trim$(mProps.CodInventario)
End Property


Public Property Let Cantidad(ByVal value As Currency)
    mProps.Cantidad = value
    mProps.objIVInventario.Modificado = True   'trabaja con  la bandera modificado del papa
End Property

Public Property Get Cantidad() As Currency
    Cantidad = mProps.Cantidad
End Property

Private Sub Class_Terminate()
    Set mProps.objIVInventario = Nothing
End Sub


Public Property Get Descripcion() As String 'Propiedad  solo de lectura
    
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "IVInventario_Descripcion", MSGERR_INVALIDO
    End If
    
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, ";") > 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario_Descripcion", _
            "El caracter ';' no es válido para la descripción."
    End If
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, "|") > 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario_Descripcion", _
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
        Err.Raise ERR_INVALIDO, "IVInventario_Unidad", MSGERR_INVALIDO
    End If
    
    mProps.Unidad = value
    mbooModificado = True
End Property

 'key puede ser IdInventario  o CodInventario
Public Sub Recuperar(key As Variant)
   Dim sql As String, i As Integer
    Dim rs As Recordset
    Dim cond As String
    cond = " IV.CodInventario='" & key & "'"
    If VarType(key) = vbString Then
        sql = "Select ivg6.codgrupo6,IV.IdInventario,  IV.CodInventario, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
        "FROM IVInventario IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad Left Join IVGrupo6 ivg6 on ivg6.idgrupo6 = iv.idgrupo6 WHERE" & cond
    Else
        sql = "Select ivg6.codgrupo6,IV.IdInventario,  IV.CodInventario, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
        "FROM IVInventario IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad Left Join IVGrupo6 ivg6 on ivg6.idgrupo6 = iv.idgrupo6 WHERE " & cond
    End If
    Set rs = Me.IVInventario.Empresa.OpenRecordset(sql)
    If rs.EOF Then 'AUC cambiado antes estuvo NOT (RS.EOF)
       Err.Raise ERR_NOHAYCODIGO, "IvFamiliaDetalle", MSGERR_NOHAYCODIGO
    Else
        mProps.IdInventario = rs!IdInventario
        mProps.CodInventario = rs!CodInventario
        mProps.Descripcion = rs!Descripcion
        If Not IsNull(rs!CodUnidad) Then mProps.Unidad = rs!CodUnidad
        If Not IsNull(rs!CostoUltimoIngreso) Then mProps.CostoUltimoIngreso = rs!CostoUltimoIngreso 'AUC 22/12/05
        If Not IsNull(rs!CodGrupo6) Then mProps.CodGrupo6 = rs!CodGrupo6
 End If
End Sub
'AUC 22/12/05
Public Property Let CostoUltimoIngreso(ByVal value As Currency)
    mProps.CostoUltimoIngreso = value
    mProps.objIVInventario.Modificado = True   'trabaja con  la bandera modificado del papa
End Property
'AUC 22/12/05
Public Property Get CostoUltimoIngreso() As Currency
    CostoUltimoIngreso = mProps.CostoUltimoIngreso
End Property
'AUC 22/12/05
Public Property Let TotalReceta(ByVal value As Currency)
    mProps.TotalReceta = value
    mProps.objIVInventario.Modificado = True   'trabaja con  la bandera modificado del papa
End Property
'AUC 22/12/05
Public Property Get TotalReceta() As Currency
    TotalReceta = mProps.TotalReceta
End Property

Public Sub RecuperarPreparacion(key As Variant, Optional SiPreparacion As String)
    Dim sql As String, i As Integer
    Dim rs As Recordset
    'select IdInventario,  CodInventario, Descripcion, Unidad   from IVInventario
    If VarType(key) = vbString Then
        sql = "Select IV.IdInventario,  IV.CodInventario, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
        " FROM IVInventario IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad WHERE CodInventario='" & key & "'"
    Else
        sql = "Select IV.IdInventario,  IV.CodInventario, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
        " FROM IVInventario IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad  WHERE IdInventario='" & key & "'"
    End If
    Set rs = Me.IVInventario.Empresa.OpenRecordset(sql)
    If rs.EOF Then
        Err.Raise ERR_NOHAYCODIGO, "IvFamiliaDetalle", MSGERR_NOHAYCODIGO
    Else
        mProps.IdInventario = rs!IdInventario
        mProps.CodInventario = rs!CodInventario
        mProps.Descripcion = rs!Descripcion
        If Not IsNull(rs!CodUnidad) Then mProps.Unidad = rs!CodUnidad
        If IVInventario.Empresa.GNOpcion.ObtenerValor("itemPreparacion") = 0 Then 'AUC toma el p.costo
            If Not IsNull(rs!CostoUltimoIngreso) Then mProps.CostoUltimoIngreso = rs!CostoUltimoIngreso 'AUC 22/12/05
        ElseIf IVInventario.Empresa.GNOpcion.ObtenerValor("itemPreparacion") = 1 Then 'AUC toma el p.venta
            mProps.pVenta = ObtienePrecio(mProps.CodInventario)
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
   mProps.objIVInventario.Modificado = True   'trabaja con  la bandera modificado del papa
End Property

'AUC 23/03/06
Public Property Get pVenta() As Currency
    pVenta = mProps.pVenta
End Property

Public Function ObtienePrecio(cod As String) As Currency
Dim item As IVInventario
Dim ix As Integer
Set item = IVInventario.Empresa.RecuperaIVInventario(cod)
With item
    If Len(IVInventario.Empresa.GNOpcion.ObtenerValor("Precio")) > 0 Then
        Select Case IVInventario.Empresa.GNOpcion.ObtenerValor("Precio")
            Case "P1"
                ix = 1
            Case "P2"
                ix = 2
            Case "P3"
                ix = 3
            Case "P4"
                ix = 4
        End Select
    Else
        ix = 1
    End If
     ObtienePrecio = .Precio(ix) + (.Precio(ix) * .PorcentajeIVA) '*  para sacar el precio MAS iva
End With
Set item = Nothing
End Function
'key puede ser CodInventario o Descripcion

Public Sub RecuperarPreparacionDescripcion(key As Variant)
   Dim sql As String, i As Integer
    Dim rs As Recordset
    Dim cond As String
        cond = " IV.descripcion='" & key & "'"
'    If IVInventario.Empresa.GNOpcion.ObtenerValor("itemPreparacion") = 0 Then 'AUC toma el p.costo
       If VarType(key) = vbString Then
            sql = "Select IV.IdInventario,  IV.CodInventario, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
            "FROM IVInventario IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad WHERE" & cond
        Else
            sql = "Select IV.IdInventario,  IV.CodInventario, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
            "FROM IVInventario IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad  WHERE " & cond
        End If
'    ElseIf IVInventario.Empresa.GNOpcion.ObtenerValor("itemPreparacion") = 1 Then 'AUC toma el p.venta
'        If VarType(key) = vbString Then
'            sql = "Select IV.IdInventario,  IV.CodInventario, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
'            "FROM IVInventario IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad WHERE" & cond
'        Else
'            sql = "Select IV.IdInventario,  IV.CodInventario, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
'            "FROM IVInventario IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad  WHERE " & cond
'        End If
'    End If
   Set rs = Me.IVInventario.Empresa.OpenRecordset(sql)
    If rs.EOF Then 'AUC cambiado antes estuvo NOT (RS.EOF)
       Err.Raise ERR_NOHAYCODIGO, "IvFamiliaDetalle", MSGERR_NOHAYCODIGO
    Else
        mProps.IdInventario = rs!IdInventario
        mProps.CodInventario = rs!CodInventario
        mProps.Descripcion = rs!Descripcion
        If Not IsNull(rs!CodUnidad) Then mProps.Unidad = rs!CodUnidad
'        If IVInventario.Empresa.GNOpcion.ObtenerValor("itemPreparacion") = 0 Then 'AUC toma el p.costo
           If Not IsNull(rs!CostoUltimoIngreso) Then mProps.CostoUltimoIngreso = rs!CostoUltimoIngreso 'AUC 22/12/05
 '       ElseIf IVInventario.Empresa.GNOpcion.ObtenerValor("itemPreparacion") = 1 Then 'AUC toma el p.venta
  '          mProps.pVenta = ObtienePrecio(mProps.CodInventario)
   '     End If
  End If
End Sub

'AUC 26/09/07 para promociones
Public Property Let xCuanto(ByVal value As Currency)
    mProps.xCuanto = value
    mProps.objIVInventario.Modificado = True   'trabaja con  la bandera modificado del papa
End Property

Public Property Get xCuanto() As Currency
    xCuanto = mProps.xCuanto
End Property

'AUC 26/09/07
Public Sub RecuperarPromocion(key As Variant, Optional SiPreparacion As String)
    Dim sql As String, i As Integer
    Dim rs As Recordset
    'select IdInventario,  CodInventario, Descripcion, Unidad   from IVInventario
    If VarType(key) = vbString Then
        sql = "Select IV.IdInventario,  IV.CodInventario, IV.Descripcion, U.CodUnidad " & _
        " FROM IVInventario IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad WHERE CodInventario='" & key & "'"
    Else
        sql = "Select IV.IdInventario,  IV.CodInventario, IV.Descripcion, U.CodUnidad " & _
        " FROM IVInventario IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad  WHERE IdInventario='" & key & "'"
    End If
    Set rs = Me.IVInventario.Empresa.OpenRecordset(sql)
    If rs.EOF Then
        Err.Raise ERR_NOHAYCODIGO, "IvFamiliaDetalle", MSGERR_NOHAYCODIGO
    Else
        mProps.IdInventario = rs!IdInventario
        mProps.CodInventario = rs!CodInventario
        mProps.Descripcion = rs!Descripcion
        mProps.Cantidad = rs!Cantidad
        mProps.xCuanto = rs!xCuanto
        mProps.Precio = rs!Precio
    End If
End Sub
'AUC 26/09/07 para promociones
Public Property Let Precio(ByVal value As Currency)
    mProps.Precio = value
    mProps.objIVInventario.Modificado = True   'trabaja con  la bandera modificado del papa
End Property
Public Property Get Precio() As Currency
    Precio = mProps.Precio
End Property

' 09/05/06  Agregado para la restriccion en la formula
Public Property Let BandPrincipal(ByVal value As Boolean)
    mProps.BandPrincipal = value
    mProps.objIVInventario.Modificado = True
End Property

Public Property Get BandPrincipal() As Boolean
    BandPrincipal = mProps.BandPrincipal
End Property


' 09/05/06  Agregado para la restriccion en la formula
Public Property Let BandModificar(ByVal value As Boolean)
    mProps.BandModificar = value
    mProps.objIVInventario.Modificado = True
End Property

Public Property Get BandModificar() As Boolean
    BandModificar = mProps.BandModificar
End Property

Public Property Let TarifaJornal(ByVal value As Currency)
    mProps.TarifaJornal = value
    mProps.objIVInventario.Modificado = True   'trabaja con  la bandera modificado del papa
End Property

Public Property Get TarifaJornal() As Currency
    TarifaJornal = mProps.TarifaJornal
End Property

Public Property Let Rendimiento(ByVal value As Currency)
    mProps.Rendimiento = value
    mProps.objIVInventario.Modificado = True   'trabaja con  la bandera modificado del papa
End Property

Public Property Get Rendimiento() As Currency
    Rendimiento = mProps.Rendimiento
End Property

Public Property Let Orden(ByVal value As Integer)
    mProps.Orden = value
    mProps.objIVInventario.Modificado = True
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property

Public Property Let Formula(ByVal value As String)
    If Len(value) > 350 Then
        Err.Raise ERR_INVALIDO, "IVInventario_Formula", MSGERR_INVALIDO
    End If
    mProps.Formula = value
    mProps.objIVInventario.Modificado = True
End Property

Public Property Get Formula() As String
    Formula = Trim$(mProps.Formula)
End Property


Public Sub RecuperarEnSucursal(key As Variant, ByVal bd As String)
End Sub

Public Sub RecuperarEnSucursales(key As Variant, ByVal bd As String, ByVal servidor As String)
  Dim sql As String, i As Integer
    Dim rs As Recordset
    Dim cond As String
    cond = " IV.CodInventario='" & key & "'"
    If VarType(key) = vbString Then
        sql = "Select IV.IdInventario,  IV.CodInventario, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
        " FROM " & bd & ".dbo.IVInventario IV LEFT JOIN " & bd & ".dbo.IVUnidad U ON U.idUnidad = IV.IdUnidad WHERE" & cond
    Else
        sql = "Select IV.IdInventario,  IV.CodInventario, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
        " FROM " & bd & ".dbo.IVInventario IV LEFT JOIN " & bd & ".dbo.IVUnidad U ON U.idUnidad = IV.IdUnidad  WHERE " & cond
    End If
    Set rs = Me.IVInventario.Empresa.OpenRecordsetParaEditSuc(sql, bd, servidor)
   If rs.EOF Then 'AUC cambiado antes estuvo NOT (RS.EOF)
       Err.Raise ERR_NOHAYCODIGO, "IvFamiliaDetalle", MSGERR_NOHAYCODIGO
    Else
        mProps.IdInventario = rs!IdInventario
        mProps.CodInventario = rs!CodInventario
        mProps.Descripcion = rs!Descripcion
        If Not IsNull(rs!CodUnidad) Then mProps.Unidad = rs!CodUnidad
        If Not IsNull(rs!CostoUltimoIngreso) Then mProps.CostoUltimoIngreso = rs!CostoUltimoIngreso 'AUC 22/12/05
  End If
End Sub

Public Property Let Tipo(ByVal value As Integer)
    mProps.Tipo = value
    mProps.objIVInventario.Modificado = True
End Property

Public Property Get Tipo() As Integer
    Tipo = mProps.Tipo
End Property

Public Property Let CostoReferencial(ByVal value As Currency)
    If mProps.CostoReferencial <> value Then
        mProps.CostoReferencial = value
        mbooModificado = True
    End If
End Property
'AUC
Public Property Get CostoReferencial() As Currency
    CostoReferencial = mProps.CostoReferencial
End Property

Public Property Let idproceso(ByVal value As Long)
 '   mProps.IdProceso = value
  '  mProps.objIVInventario.Modificado = True
End Property

Public Property Get idproceso() As Long
    idproceso = mProps.idproceso
End Property

Public Property Let CodProceso(ByVal value As String)
Dim rs As Recordset, sql As String
    sql = "Select idproceso from ivproceso where codproceso = '" & value & "'"
    Set rs = Me.IVInventario.Empresa.OpenRecordset(sql)
    If Not rs.EOF Then
        mProps.idproceso = rs!idproceso
    Else
        mProps.idproceso = 0
    End If
    mProps.CodProceso = value
    mProps.objIVInventario.Modificado = True
End Property

Public Property Get CodProceso() As String
    CodProceso = Trim$(mProps.CodProceso)
End Property

Friend Function Clone() As IVFamiliaDetalle
    Dim nuevo As IVFamiliaDetalle
    
    Set nuevo = New IVFamiliaDetalle
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)          'Para el método Clone
    mProps = value
'    mProps.objIVInventario.Modificado (True)
End Property

Public Property Let CodGrupo6(ByVal value As String)
    mProps.CodGrupo6 = value
    'mProps.objIVInventario.Modificado = True
End Property
Public Property Get CodGrupo6() As String
    CodGrupo6 = mProps.CodGrupo6
End Property

