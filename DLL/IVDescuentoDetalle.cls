VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVDescuentoDetallePC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***Diego 15/09/2003
Private Type T_PROPS
    IdDescuento As Long
    IdDescuentoDetalle As Long
    Valor As Currency
    NumPCGrupo  As Byte
    IDPCGrupo  As Integer
    CodPCGrupo As String      'revisar  si nos sirve
    
    NumIVGrupo  As Byte
    IDIVGrupo  As Integer
    CodIVGrupo As String      'revisar  si nos sirve
    
    
    
    Orden As Long
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
            mProps.IdDescuentoDetalle = !IdDescuentoDetalle
            If Not IsNull(!NumPCGrupo) Then mProps.NumPCGrupo = !NumPCGrupo
            If Not IsNull(!IDPCGrupo) Then mProps.IDPCGrupo = !IDPCGrupo
            If Not IsNull(!NumIVGrupo) Then mProps.NumIVGrupo = !NumIVGrupo
            If Not IsNull(!IDIVGrupo) Then mProps.IDIVGrupo = !IDIVGrupo

            mProps.Valor = !Valor
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            'recupera  el codigo  de PCGrupo
            On Error Resume Next
            If mProps.NumPCGrupo > 0 And mProps.NumPCGrupo < PCGRUPO_MAX Then  'Abre todo
                sql = "Select CodGrupo" & mProps.NumPCGrupo & " From PCGrupo" & mProps.NumPCGrupo & _
                      " Where IDGrupo" & mProps.NumPCGrupo & " = " & mProps.IDPCGrupo
                Set rs1 = mProps.objIVDescuento.Empresa.OpenRecordset(sql)
                If Not rs1.EOF Then
                    mProps.CodPCGrupo = rs1.Fields(0).value
                End If
                rs1.Close
                Set rs1 = Nothing
            End If
            If mProps.NumIVGrupo > 0 And mProps.NumIVGrupo < IVGRUPO_MAX Then  'Abre todo
                sql = "Select CodGrupo" & mProps.NumIVGrupo & " From IVGrupo" & mProps.NumIVGrupo & _
                      " Where IDGrupo" & mProps.NumIVGrupo & " = " & mProps.IDIVGrupo
                Set rs1 = mProps.objIVDescuento.Empresa.OpenRecordset(sql)
                If Not rs1.EOF Then
                    mProps.CodIVGrupo = rs1.Fields(0).value
                End If
                rs1.Close
                Set rs1 = Nothing
            End If
        
        End If
    End With
End Sub

Public Function Clone() As IVDescuentoDetalle
    Dim nuevo As IVDescuentoDetalle
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New IVDescuentoDetalle
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

Public Property Let Valor(ByVal value As Currency)
    mProps.Valor = value
    mProps.objIVDescuento.Modificado = True
End Property

Public Property Get Valor() As Currency
    Valor = mProps.Valor
End Property


Public Property Let NumPCGrupo(ByVal value As Byte)
    mProps.NumPCGrupo = value
    mProps.objIVDescuento.Modificado = True
End Property

Public Property Get NumPCGrupo() As Byte
    NumPCGrupo = mProps.NumPCGrupo
End Property



Public Property Get CodPCGrupo() As String
    CodPCGrupo = mProps.CodPCGrupo
End Property


Public Property Let CodPCGrupo(value As String)
    Dim rs As Recordset, sql As String
    mProps.CodPCGrupo = value
    If value = "" Then
        mProps.IDPCGrupo = 0
    Else
        'actualiza  el IDGrupo
        sql = "Select IdGrupo" & mProps.NumPCGrupo & " From PCGrupo" & mProps.NumPCGrupo & _
              " Where CodGrupo" & mProps.NumPCGrupo & " = '" & value & "'"
        Set rs = mProps.objIVDescuento.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then
            mProps.IDPCGrupo = rs.Fields(0).value
        End If
        rs.Close
        Set rs = Nothing
    End If
    mProps.objIVDescuento.Modificado = True
End Property


Public Property Get IDPCGrupo() As Long
    IDPCGrupo = mProps.IDPCGrupo
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




