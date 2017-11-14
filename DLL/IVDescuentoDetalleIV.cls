VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVDescuentoDetalleIV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***Diego 15/09/2003
Private Type T_PROPS
    IdDescuento As Long
    IdDescuentoDetalleIV As Long

    NumIVGrupo  As Byte
    IDIVGrupo  As Integer
    CodIVGrupo As String      'revisar  si nos sirve
    idInventario As Long
    CodInventario As String      'revisar  si nos sirve
    Descuento As Currency
    Descripcion As String
    Unidad As String
    Orden As Long
    CantMin As Currency 'AUC agregado para italdeli
    objIVDescuento As IVDescuento   'Referencia al Objeto Padre
End Type

Private mProps As T_PROPS
'Private mbooEsNuevo As Boolean
'Private mbooModificado As Boolean


'Establece relacion  con el PADRE
Friend Property Set IVDescuento(objIVDescuento As IVDescuento)
    Set mProps.objIVDescuento = objIVDescuento
End Property

Public Property Get IVDescuento() As IVDescuento
    Set IVDescuento = mProps.objIVDescuento
End Property

Friend Sub Recuperar(rs As Recordset)
    Dim sql As String, rs1  As Recordset
    With rs
        If Not .EOF Then
            mProps.IdDescuento = !IdDescuento
            mProps.IdDescuentoDetalleIV = !IdDescuentoDetalleIV
            If Not IsNull(!NumIVGrupo) Then mProps.NumIVGrupo = !NumIVGrupo
            If Not IsNull(!IDIVGrupo) Then mProps.IDIVGrupo = !IDIVGrupo
            If Not IsNull(!idInventario) Then mProps.idInventario = !idInventario
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            If Not IsNull(!Descuento) Then mProps.Descuento = !Descuento
            If Not IsNull(!CantMin) Then mProps.CantMin = !CantMin
            'recupera  el codigo  de PCGrupo
            On Error Resume Next
            If mProps.NumIVGrupo > 0 And mProps.NumIVGrupo < IVGRUPO_MAX Then  'Abre todo
                sql = "Select CodGrupo" & mProps.NumIVGrupo & " From IVGrupo" & mProps.NumIVGrupo & _
                      " Where IDGrupo" & mProps.NumIVGrupo & " = " & mProps.IDIVGrupo
                Set rs1 = mProps.objIVDescuento.Empresa.OpenRecordset(sql)
                If Not rs1.EOF Then
                    mProps.CodIVGrupo = rs1.Fields(0).value
                End If
                rs1.Close
                Set rs1 = Nothing
                'CARGA DATOS ITEMS TAMBIEN
                sql = "Select CodInventario, IVINVENTARIO.Descripcion, IVU.codunidad From IvInventario "
                sql = sql & "LEFT JOIN IVUNIDAD IVU on IVU.IDUNIDAD = IVINVENTARIO.IdUNIDAD"
                sql = sql & " Where IDInventario = " & mProps.idInventario
                Set rs1 = mProps.objIVDescuento.Empresa.OpenRecordset(sql)
                If Not rs1.EOF Then
                    mProps.CodInventario = rs1.Fields(0).value
                    mProps.Descripcion = rs1.Fields(1).value
                    mProps.Unidad = rs1.Fields(2).value
                End If
                rs1.Close
                Set rs1 = Nothing
            Else
                sql = "Select CodInventario, IVINVENTARIO.Descripcion, IVU.codunidad From IvInventario "
                sql = sql & "LEFT JOIN IVUNIDAD IVU on IVU.IDUNIDAD = IVINVENTARIO.IdUNIDAD"
                sql = sql & " Where IDInventario = " & mProps.idInventario
                Set rs1 = mProps.objIVDescuento.Empresa.OpenRecordset(sql)
                If Not rs1.EOF Then
                    mProps.CodInventario = rs1.Fields(0).value
                    mProps.Descripcion = rs1.Fields(1).value
                    mProps.Unidad = rs1.Fields(2).value
                End If
                rs1.Close
                Set rs1 = Nothing
            End If

        End If
    End With
End Sub

Public Function Clone() As IVDescuentoDetalleIV
    Dim nuevo As IVDescuentoDetalleIV

    'Copia las propiedades al nuevo objeto
    Set nuevo = New IVDescuentoDetalleIV
    nuevo.Props = mProps

    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdDescuento = 0
    'mbooModificado = True
End Property


Public Property Get IdDescuento() As Long
    IdDescuento = mProps.IdDescuento
End Property

Public Property Let IdDescuento(ByVal value As Long)
    mProps.IdDescuento = value
End Property

Public Property Get Orden() As Long
    Orden = mProps.Orden
End Property

Public Property Let Orden(ByVal value As Long)
    mProps.Orden = value
End Property



Private Sub Class_Terminate()
    Set mProps.objIVDescuento = Nothing
End Sub


Public Property Let NumIVGrupo(ByVal value As Byte)
    mProps.NumIVGrupo = value
    mProps.objIVDescuento.Modificado = True
End Property

Public Property Get NumIVGrupo() As Byte
    NumIVGrupo = mProps.NumIVGrupo
End Property



Public Property Get CodIVGrupo() As String
    CodIVGrupo = mProps.CodIVGrupo
End Property


Public Property Let CodIVGrupo(value As String)
    Dim rs As Recordset, sql As String
    mProps.CodIVGrupo = value
    If value = "" Then
        mProps.IDIVGrupo = 0
    Else
        'actualiza  el IDGrupo
        sql = "Select IdGrupo" & mProps.NumIVGrupo & " From IVGrupo" & mProps.NumIVGrupo & _
              " Where CodGrupo" & mProps.NumIVGrupo & " = '" & value & "'"
        Set rs = mProps.objIVDescuento.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then
            mProps.IDIVGrupo = rs.Fields(0).value
        End If
        rs.Close
        Set rs = Nothing
    End If
    mProps.objIVDescuento.Modificado = True
End Property

Public Property Get IDIVGrupo() As Long
    IDIVGrupo = mProps.IDIVGrupo
End Property

Public Property Get CodInventario() As String
    CodInventario = mProps.CodInventario
End Property

Public Property Let CodInventario(value As String)
    Dim rs As Recordset, sql As String
    mProps.CodInventario = value
    If value = "" Then
        mProps.idInventario = 0
    Else
        'actualiza  el IDGrupo
        sql = "Select IdInventario From IvInventario"
        sql = sql & " Where CodInventario = '" & value & "'"
        Set rs = mProps.objIVDescuento.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then
            mProps.idInventario = rs.Fields(0).value
        End If
        rs.Close
        Set rs = Nothing
    End If
    mProps.objIVDescuento.Modificado = True
End Property

Public Property Get idInventario() As Long
    idInventario = mProps.idInventario
End Property

Public Property Let Descuento(ByVal value As Currency)
    mProps.Descuento = value
    mProps.objIVDescuento.Modificado = True
End Property

Public Property Get Descuento() As Currency
    Descuento = mProps.Descuento
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let Descripcion(ByVal value As String)
    mProps.Descripcion = value
    mProps.objIVDescuento.Modificado = True
End Property


Public Sub RecuperarDescripcion(key As Variant)
   Dim sql As String, i As Integer
    Dim rs As Recordset
    Dim cond As String
        cond = " Descripcion ='" & key & "'"
       If VarType(key) = vbString Then
            sql = "Select idInventario,codInventario,Descripcion " & _
            "FROM IvInventrio WHERE" & cond
        Else
            sql = "Select idInventario,codInventario,Descripcion " & _
            "FROM IvInventario WHERE" & cond
        End If
   Set rs = Me.IVDescuento.Empresa.OpenRecordset(sql)
    If rs.EOF Then 'AUC cambiado antes estuvo NOT (RS.EOF)
       Err.Raise ERR_NOHAYCODIGO, "IvFamiliaDetalle", MSGERR_NOHAYCODIGO
    Else
            mProps.idInventario = rs!idInventario
            mProps.CodInventario = rs!CodInventario
            mProps.Descripcion = rs!Descripcion
  End If
End Sub



Public Sub recuperarNEW(key As Variant)
   Dim sql As String, i As Integer
    Dim rs As Recordset
    Dim cond As String
    
        cond = " CodInventario='" & key & "'"
        If VarType(key) = vbString Then
            sql = "Select idInventario, CodInventario, descripcion "
            sql = sql & " FROM IvInventario WHERE" & cond
        Else
            sql = "Select idInventario, CodInventario, descripcion"
            sql = sql & " FROM IvInventario WHERE" & cond
        End If
    Set rs = Me.IVDescuento.Empresa.OpenRecordset(sql)
    If rs.EOF Then 'AUC cambiado antes estuvo NOT (RS.EOF)
       Err.Raise ERR_NOHAYCODIGO, "IVDescuentoDetalleIV", MSGERR_NOHAYCODIGO
    Else
                mProps.idInventario = rs!idInventario
                mProps.CodInventario = rs!CodInventario
                mProps.Descripcion = rs!Descripcion


  End If
End Sub

Public Sub RecuperarNombre(key As Variant)
   Dim sql As String, i As Integer
    Dim rs As Recordset
    Dim cond As String
        cond = " Descripcion ='" & key & "'"
        
        If VarType(key) = vbString Then
            sql = "Select idInventario, CodInventario, descripcion "
            sql = sql & " FROM IvInventario WHERE" & cond
        Else
            sql = "Select idInventario, CodInventario, descripcion"
            sql = sql & " FROM IvInventario WHERE" & cond
        End If
        
        
   Set rs = Me.IVDescuento.Empresa.OpenRecordset(sql)
    If rs.EOF Then 'AUC cambiado antes estuvo NOT (RS.EOF)
       Err.Raise ERR_NOHAYCODIGO, "IvFamiliaDetalle", MSGERR_NOHAYCODIGO
    Else
                mProps.idInventario = rs!idInventario
                mProps.CodInventario = rs!CodInventario
                mProps.Descripcion = rs!Descripcion
  End If
End Sub



'AUC agrega unidad MEGALIMPIO
Friend Sub RecuperarUnidad(rs As Recordset)
    Dim sql As String, rs1  As Recordset
    With rs
        If Not .EOF Then
            mProps.IdDescuento = !IdDescuento
            mProps.IdDescuentoDetalleIV = !IdDescuentoDetalleIV
            If Not IsNull(!NumIVGrupo) Then mProps.NumIVGrupo = !NumIVGrupo
            If Not IsNull(!IDIVGrupo) Then mProps.IDIVGrupo = !IDIVGrupo
            If Not IsNull(!idInventario) Then mProps.idInventario = !idInventario
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            If Not IsNull(!Descuento) Then mProps.Descuento = !Descuento
            'recupera  el codigo  de PCGrupo
            On Error Resume Next
            If mProps.NumIVGrupo > 0 And mProps.NumIVGrupo < IVGRUPO_MAX Then  'Abre todo
                sql = "Select CodGrupo" & mProps.NumIVGrupo & " From IVGrupo" & mProps.NumIVGrupo & _
                      " Where IDGrupo" & mProps.NumIVGrupo & " = " & mProps.IDIVGrupo
                Set rs1 = mProps.objIVDescuento.Empresa.OpenRecordset(sql)
                If Not rs1.EOF Then
                    mProps.CodIVGrupo = rs1.Fields(0).value
                End If
                rs1.Close
                Set rs1 = Nothing
                'CARGA DATOS ITEMS TAMBIEN
                sql = "Select CodInventario, IVINVENTARIO.Descripcion, IVU.codunidad From IvInventario "
                sql = sql & "LEFT JOIN IVUNIDAD IVU on IVU.IDUNIDAD = IVINVENTARIO.IdUNIDAD"
               sql = sql & " Where IDInventario = " & mProps.idInventario
                Set rs1 = mProps.objIVDescuento.Empresa.OpenRecordset(sql)
                If Not rs1.EOF Then
                    mProps.CodInventario = rs1.Fields(0).value
                    mProps.Descripcion = rs1.Fields(1).value
                    mProps.Unidad = rs1.Fields(2).value
               End If
                rs1.Close
                Set rs1 = Nothing
            Else
                sql = "Select CodInventario, IVINVENTARIO.Descripcion, IVU.codunidad From IvInventario "
                sql = sql & "LEFT JOIN IVUNIDAD IVU on IVU.IDUNIDAD = IVINVENTARIO.IdUNIDAD"
               sql = sql & " Where IDInventario = " & mProps.idInventario
                Set rs1 = mProps.objIVDescuento.Empresa.OpenRecordset(sql)
                If Not rs1.EOF Then
                    mProps.CodInventario = rs1.Fields(0).value
                    mProps.Descripcion = rs1.Fields(1).value
                    mProps.Unidad = rs1.Fields(2).value
               End If
                rs1.Close
                Set rs1 = Nothing
            End If
        End If
    End With
End Sub

Public Property Get Unidad() As String
    Unidad = mProps.Unidad
End Property

Public Property Let Unidad(ByVal value As String)
    mProps.Unidad = value
    mProps.objIVDescuento.Modificado = True
End Property



Public Property Let CantMin(ByVal value As Currency)
    mProps.CantMin = value
    mProps.objIVDescuento.Modificado = True
End Property

Public Property Get CantMin() As Currency
    CantMin = mProps.CantMin
End Property

