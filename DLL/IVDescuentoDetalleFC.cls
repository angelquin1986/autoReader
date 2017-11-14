VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVDescuentoDetalleFC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***Diego 15/09/2003
Private Type T_PROPS
    IdDescuento As Long
    IdDescuentoDetalle As Long
    Descuento As Currency
    Penal  As Currency
    IDForma  As Integer
    CodForma As String      'revisar  si nos sirve

    Orden As Long
    objIVDescuento As IVDescuento   'Referencia al Objeto Padre
End Type

Private mProps As T_PROPS


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
            mProps.IdDescuentoDetalle = !IdDescuentoDetallefc
            
            If Not IsNull(!IDForma) Then mProps.IDForma = !IDForma
            mProps.Descuento = !Descuento
            mProps.Penal = !Penal
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            'recupera  el codigo  de Forma
            On Error Resume Next
            
                sql = "Select CodForma From TSFormaCobroPago"
                sql = sql & " Where IDForma = " & mProps.IDForma
                Set rs1 = mProps.objIVDescuento.Empresa.OpenRecordset(sql)
                If Not rs1.EOF Then
                    mProps.CodForma = rs1.Fields(0).value
                End If
                rs1.Close
                Set rs1 = Nothing


        End If
    End With
End Sub

Public Function Clone() As IVDescuentoDetalleFC
    Dim nuevo As IVDescuentoDetalleFC

    'Copia las propiedades al nuevo objeto
    Set nuevo = New IVDescuentoDetalleFC
    nuevo.Props = mProps

    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdDescuento = 0
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

Public Property Let Descuento(ByVal value As Currency)
    mProps.Descuento = value
    mProps.objIVDescuento.Modificado = True
End Property

Public Property Get Descuento() As Currency
    Descuento = mProps.Descuento
End Property


Public Property Let Penal(ByVal value As Currency)
    mProps.Penal = value
    mProps.objIVDescuento.Modificado = True
End Property

Public Property Get Penal() As Currency
    Penal = mProps.Penal
End Property

Public Property Get CodForma() As String
    CodForma = mProps.CodForma
End Property


Public Property Let CodForma(value As String)
    Dim rs As Recordset, sql As String
    mProps.CodForma = value
    If value = "" Then
        mProps.IDForma = 0
    Else
        'actualiza  el IDGrupo
        sql = "Select IdForma  From TSFormaCobroPago "
        sql = sql & " Where CodForma  = '" & value & "'"
        Set rs = mProps.objIVDescuento.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then
            mProps.IDForma = rs.Fields(0).value
        End If
        rs.Close
        Set rs = Nothing
    End If
    mProps.objIVDescuento.Modificado = True
End Property


Public Property Get IDForma() As Long
    IDForma = mProps.IDForma
End Property



Private Sub Class_Terminate()
    Set mProps.objIVDescuento = Nothing
End Sub


