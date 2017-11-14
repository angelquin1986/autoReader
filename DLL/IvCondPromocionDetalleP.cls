VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IvCondPromocionDetalleP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***Diego 15/09/2003
Private Type T_PROPS
    IdPromocion As Long
    IdCondPromocionDetalleP As Long
    Idinventario As Integer
    CodInventario As String      'revisar  si nos sirve
    Descripcion As String      'revisar  si nos sirve
    
    cant As Currency
    PrecioU As Currency
    Orden As Long
    objIVPromocion As IVPromocion   'Referencia al Objeto Padre
End Type

Private mProps As T_PROPS


'Establece relacion  con el PADRE
Friend Property Set IVPromocion(objIVPromocion As IVPromocion)
    Set mProps.objIVPromocion = objIVPromocion
End Property

Public Property Get IVPromocion() As IVPromocion
    Set IVPromocion = mProps.objIVPromocion
End Property


Public Function Clone() As IvCondPromocionDetalleP
    Dim nuevo As IvCondPromocionDetalleP

    'Copia las propiedades al nuevo objeto
    Set nuevo = New IvCondPromocionDetalleP
    nuevo.Props = mProps

    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdPromocion = 0
    'mbooModificado = True
End Property


Public Property Get IdPromocion() As Long
    IdPromocion = mProps.IdPromocion
End Property

Public Property Let IdPromocion(ByVal value As Long)
    mProps.IdPromocion = value
End Property


Public Property Get CodInventario() As String
    CodInventario = mProps.CodInventario
End Property

Public Property Let CodInventario(value As String)
    Dim rs As Recordset, sql As String
    mProps.CodInventario = value
    If value = "" Then
        mProps.Idinventario = 0
    Else
        'actualiza  el IDGrupo
        sql = "Select IdInventario From IvInventario"
        sql = sql & " Where CodInventario = '" & value & "'"
        Set rs = mProps.objIVPromocion.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then
            mProps.Idinventario = rs.Fields(0).value
        End If
        rs.Close
        Set rs = Nothing
    End If
    mProps.objIVPromocion.Modificado = True
End Property

Public Property Get Idinventario() As Long
    Idinventario = mProps.Idinventario
End Property


Public Property Get Orden() As Long
    Orden = mProps.Orden
End Property

Public Property Let Orden(ByVal value As Long)
    mProps.Orden = value
End Property



Private Sub Class_Initialize()
    mProps.cant = 0
    mProps.PrecioU = 0
   
End Sub

Private Sub Class_Terminate()
    Set mProps.objIVPromocion = Nothing
End Sub




Public Property Let cant(ByVal value As Currency)
    mProps.cant = value
    mProps.objIVPromocion.Modificado = True
End Property

Public Property Get cant() As Currency
    cant = mProps.cant
End Property

Public Property Let PrecioU(ByVal value As Currency)
    mProps.PrecioU = value
    mProps.objIVPromocion.Modificado = True
End Property

Public Property Get PrecioU() As Currency
    PrecioU = mProps.PrecioU
End Property

Public Sub Recuperar(key As Variant)
   Dim sql As String, i As Integer
    Dim rs As Recordset
    Dim cond As String
    
        cond = " IV.CodInventario='" & key & "'"
        If VarType(key) = vbString Then
            sql = "Select IV.IdInventario,  IV.CodInventario, IV.Descripcion, U.CodUnidad "
            sql = sql & " FROM IVInventario IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad WHERE" & cond
        Else
            sql = "Select IV.IdInventario,  IV.CodInventario, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso "
            sql = sql & " FROM IVInventario IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad  WHERE " & cond
        End If
    Set rs = Me.IVPromocion.Empresa.OpenRecordset(sql)
    If rs.EOF Then 'AUC cambiado antes estuvo NOT (RS.EOF)
       Err.Raise ERR_NOHAYCODIGO, "IvFamiliaDetalle", MSGERR_NOHAYCODIGO
    Else
                mProps.Idinventario = rs!Idinventario
                mProps.CodInventario = rs!CodInventario
                mProps.Descripcion = rs!Descripcion
                mProps.cant = 0
                mProps.PrecioU = 0
  End If
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
    mProps.objIVPromocion.Modificado = True
End Property

Public Sub RecuperarDescripcion(key As Variant)
   Dim sql As String, i As Integer
    Dim rs As Recordset
    Dim cond As String
        cond = " IV.descripcion='" & key & "'"
       If VarType(key) = vbString Then
            sql = "Select IV.IdInventario,  IV.CodInventario, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
            "FROM IVInventario IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad WHERE" & cond
        Else
            sql = "Select IV.IdInventario,  IV.CodInventario, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
            "FROM IVInventario IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad  WHERE " & cond
        End If
   Set rs = Me.IVPromocion.Empresa.OpenRecordset(sql)
    If rs.EOF Then 'AUC cambiado antes estuvo NOT (RS.EOF)
       Err.Raise ERR_NOHAYCODIGO, "IvFamiliaDetalle", MSGERR_NOHAYCODIGO
    Else
            mProps.Idinventario = rs!Idinventario
            mProps.CodInventario = rs!CodInventario
            mProps.Descripcion = rs!Descripcion
            mProps.cant = 0
            mProps.PrecioU = 0
            
  End If
End Sub




