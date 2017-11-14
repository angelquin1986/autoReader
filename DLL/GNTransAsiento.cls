VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GNTransAsiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"GNTrans"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type T_PROP
    Id As Long
    IdCuenta As Long
    CodCuenta As String
    Descripcion As String
    Formula As String
    BandDeudor As Boolean
    Comentario As String
    Orden As Integer
    
    objGNTrans As GNTrans           'Referencia al objeto padre
End Type

Private mProps As T_PROP



Friend Sub Grabar()
    Dim sql As String, rs As Recordset
    
    sql = "SELECT * FROM GNTransAsiento WHERE 1=0"
    Set rs = mProps.objGNTrans.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !CodTrans = mProps.objGNTrans.CodTrans
        !IdCuenta = mProps.IdCuenta
        !Descripcion = mProps.Descripcion
        !Formula = mProps.Formula
        !BandDeudor = mProps.BandDeudor
        !Comentario = mProps.Comentario
        !Orden = mProps.Orden
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

Friend Sub Recuperar(Id As Long)
    Dim sql As String, rs As Recordset

    mProps.Id = Id
    sql = "SELECT A.*,B.CodCuenta FROM GNTransAsiento A LEFT JOIN CTCuenta B " & _
          " ON A.IdCuenta=B.IdCuenta " & _
          "WHERE A.Id=" & mProps.Id
    Set rs = mProps.objGNTrans.Empresa.OpenRecordset(sql)
    With rs
        If Not .EOF Then
            If Not IsNull(!IdCuenta) Then mProps.IdCuenta = !IdCuenta
            If Not IsNull(!CodCuenta) Then mProps.CodCuenta = !CodCuenta
            If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
            If Not IsNull(!Formula) Then mProps.Formula = !Formula
            If Not IsNull(!BandDeudor) Then mProps.BandDeudor = !BandDeudor
            If Not IsNull(!Comentario) Then mProps.Comentario = !Comentario
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
        End If
        .Close
    End With
    Set rs = Nothing
End Sub

Friend Function Clone() As GNTransAsiento
    Dim nuevo As GNTransAsiento
    
    Set nuevo = New GNTransAsiento
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROP)          'Para el método Clone
    mProps = value
    mProps.Id = 0
    mProps.objGNTrans.SetModificado
End Property

Friend Property Set GNTrans(value As GNTrans)
    Set mProps.objGNTrans = value
End Property

Public Property Get GNTrans() As GNTrans
    Set GNTrans = mProps.objGNTrans
End Property



Public Property Get Id() As Long
    Id = mProps.Id
End Property


Public Property Let IdCuenta(ByVal value As Long)
    mProps.IdCuenta = value
    mProps.objGNTrans.SetModificado
End Property

Public Property Get IdCuenta() As Long
    IdCuenta = mProps.IdCuenta
End Property


Public Property Let CodCuenta(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCuenta Then
        If Len(value) > 0 Then
            'Actualiza IdCuenta también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objGNTrans.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNTransAsiento_CodCuenta", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuenta = rs!IdCuenta
            End If
            rs.Close
        Else
            IdCuenta = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodCuenta = value
    mProps.objGNTrans.SetModificado
End Property

Public Property Get CodCuenta() As String
    CodCuenta = Trim$(mProps.CodCuenta)
End Property



Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "GNTransAsiento_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = value
    mProps.objGNTrans.SetModificado
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let Formula(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "GNTransAsiento_Formula", MSGERR_INVALIDO
    End If
    
    mProps.Formula = value
    mProps.objGNTrans.SetModificado
End Property

Public Property Get Formula() As String
    Formula = Trim$(mProps.Formula)
End Property

Public Property Let BandDeudor(ByVal value As Boolean)
    mProps.BandDeudor = value
    mProps.objGNTrans.SetModificado
End Property

Public Property Get BandDeudor() As Boolean
    BandDeudor = mProps.BandDeudor
End Property

Public Property Let Comentario(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "GNTransAsiento_Comentario", MSGERR_INVALIDO
    End If
    
    mProps.Comentario = value
    mProps.objGNTrans.SetModificado
End Property

Public Property Get Comentario() As String
    Comentario = Trim$(mProps.Comentario)
End Property


Public Property Let Orden(ByVal value As Integer)
    mProps.Orden = value
    mProps.objGNTrans.SetModificado
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property



Private Sub Class_Terminate()
    Set mProps.objGNTrans = Nothing
End Sub


