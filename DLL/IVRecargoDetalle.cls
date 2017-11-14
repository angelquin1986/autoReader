VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVRecargoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***Diego 15/09/2003
Private Type T_PROPS
    IdRecargo As Long
    IdRecargoDetalle As Long
    Valor As Currency
    NumPCGrupo  As Byte
    IDPCGrupo  As Integer
    CodPCGrupo As String      'revisar  si nos sirve
    Orden As Long
    objIVRecargo As IVRecargo   'Referencia al Objeto Padre
End Type

Private mProps As T_PROPS
'Private mbooEsNuevo As Boolean
'Private mbooModificado As Boolean


'Establece relacion  con el PADRE
Friend Property Set IVRecargo(objIVRecargo As IVRecargo)
    Set mProps.objIVRecargo = objIVRecargo
End Property

Public Property Get IVRecargo() As IVRecargo
    Set IVRecargo = mProps.objIVRecargo
End Property

Friend Sub Recuperar(rs As Recordset)
    Dim sql As String, rs1  As Recordset
    With rs
        If Not .EOF Then
            mProps.IdRecargo = !IdRecargo
            mProps.IdRecargoDetalle = !IdRecargoDetalle
            If Not IsNull(!NumPCGrupo) Then mProps.NumPCGrupo = !NumPCGrupo
            If Not IsNull(!IDPCGrupo) Then mProps.IDPCGrupo = !IDPCGrupo
            'If Not IsNull(!CodPCGrupo) Then mProps.CodPCGrupo = !CodPCGrupo
            mProps.Valor = !Valor
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            'recupera  el codigo  de PCGrupo
            On Error Resume Next
            If mProps.NumPCGrupo > 0 And mProps.NumPCGrupo < PCGRUPO_MAX Then  'Abre todo
                sql = "Select CodGrupo" & mProps.NumPCGrupo & " From PCGrupo" & mProps.NumPCGrupo & _
                      " Where IDGrupo" & mProps.NumPCGrupo & " = " & mProps.IDPCGrupo
                Set rs1 = mProps.objIVRecargo.Empresa.OpenRecordset(sql)
                If Not rs1.EOF Then
                    mProps.CodPCGrupo = rs1.Fields(0).value
                End If
                rs1.Close
                Set rs1 = Nothing
            End If
        End If
    End With
End Sub

Public Function Clone() As IVRecargoDetalle
    Dim nuevo As IVRecargoDetalle
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New IVRecargoDetalle
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdRecargo = 0
    'mbooModificado = True
End Property


Public Property Get IdRecargo() As Long
    IdRecargo = mProps.IdRecargo
End Property

Public Property Let IdRecargo(ByVal value As Long)
    mProps.IdRecargo = value
End Property

Public Property Get Orden() As Long
    Orden = mProps.Orden
End Property

Public Property Let Orden(ByVal value As Long)
    mProps.Orden = value
End Property

Public Property Let Valor(ByVal value As Currency)
    mProps.Valor = value
    mProps.objIVRecargo.Modificado = True
End Property

Public Property Get Valor() As Currency
    Valor = mProps.Valor
End Property


Public Property Let NumPCGrupo(ByVal value As Byte)
    mProps.NumPCGrupo = value
    mProps.objIVRecargo.Modificado = True
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
        Set rs = mProps.objIVRecargo.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then
            mProps.IDPCGrupo = rs.Fields(0).value
        End If
        rs.Close
        Set rs = Nothing
    End If
    mProps.objIVRecargo.Modificado = True
End Property


Public Property Get IDPCGrupo() As Long
    IDPCGrupo = mProps.IDPCGrupo
End Property



Private Sub Class_Initialize()
    'Valores predeterminados
    'mbooEsNuevo = True
    'mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objIVRecargo = Nothing
End Sub
