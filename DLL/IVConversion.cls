VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdInventario As Long
    IdUnidadOri As Long
    IdUnidadDes As Long
    CodUnidadOri As String
    CodUnidadDes As String
    Cantidad As Currency
    CantidadDou As Double
    Orden As Integer
    Precio As Currency 'AUC 28/09/07 para promociones
    objIVInventario As IVInventario          'Referencia al objeto padre
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
    mProps.objIVInventario.Modificado = True
End Property

Friend Property Get IdInventario() As Long
    IdInventario = mProps.IdInventario
End Property

Friend Property Let IdUnidadOri(ByVal value As Long)
    mProps.IdUnidadOri = value
    mProps.objIVInventario.Modificado = True
End Property

Friend Property Get IdUnidadOri() As Long
    IdUnidadOri = mProps.IdUnidadOri
End Property
Friend Property Let IdUnidadDes(ByVal value As Long)
    mProps.IdUnidadDes = value
    mProps.objIVInventario.Modificado = True
End Property

Friend Property Get IdUnidadDes() As Long
    IdUnidadDes = mProps.IdUnidadDes
End Property

Public Property Get CodUnidadOri() As String
    CodUnidadOri = Trim$(mProps.CodUnidadOri)
End Property

Public Property Let CodUnidadOri(ByVal value As String)
  Dim sql As String, rs As Recordset
    If value <> mProps.CodUnidadOri Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdUnidad FROM IVUnidad WHERE CodUnidad='" & value & "'"
            Set rs = mProps.objIVInventario.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVInventario_CodUnidad", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdUnidadOri = rs!IdUnidad
            End If
            rs.Close
        Else
            IdUnidadOri = 0
        End If
        
        mProps.CodUnidadOri = value
        mProps.objIVInventario.Modificado = True
    End If

    Set rs = Nothing
End Property

Public Property Get CodUnidadDes() As String
    CodUnidadDes = Trim$(mProps.CodUnidadDes)
End Property

Public Property Let CodUnidadDes(ByVal value As String)
     Dim sql As String, rs As Recordset
    If value <> mProps.CodUnidadDes Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdUnidad FROM IVUnidad WHERE CodUnidad='" & value & "'"
            Set rs = mProps.objIVInventario.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVInventario_CodUnidad", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdUnidadDes = rs!IdUnidad
            End If
            rs.Close
        Else
            IdUnidadDes = 0
        End If
        mProps.CodUnidadDes = value
        mProps.objIVInventario.Modificado = True
    End If

    Set rs = Nothing
    
End Property


Public Property Let Cantidad(ByVal value As Currency)
    mProps.Cantidad = value
    mProps.objIVInventario.Modificado = True
End Property

Public Property Get Cantidad() As Currency
    Cantidad = mProps.Cantidad
End Property
Public Property Let Precio(ByVal value As Currency)
    mProps.Precio = value
    mProps.objIVInventario.Modificado = True
End Property

Public Property Get Precio() As Currency
    Precio = mProps.Precio
End Property


Private Sub Class_Terminate()
    Set mProps.objIVInventario = Nothing
End Sub

Public Property Let Orden(ByVal value As Integer)
    mProps.Orden = value
    mProps.objIVInventario.Modificado = True
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property

Public Property Let CantidadDou(ByVal value As Double)
    mProps.CantidadDou = value
    mProps.objIVInventario.Modificado = True
End Property

Public Property Get CantidadDou() As Double
    CantidadDou = mProps.CantidadDou
End Property
