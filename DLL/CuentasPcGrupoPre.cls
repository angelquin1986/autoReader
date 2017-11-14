VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CuentasPcGrupoPre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    idPcGrupo As Long
    idElemento As Long
    IdCuenta As Long
    CodCuenta As String
    FechaGrabado As Date
    objPCG As PcGrupo
End Type

Private mProps As T_PROPS
Private mbooModificado As Boolean
Private mbooEsNuevo As Boolean

Friend Property Set objPCG(ByVal obj As PcGrupo)
    Set mProps.objPCG = obj
End Property

Public Property Get objPCG() As PcGrupo
    Set objPCG = mProps.objPCG
End Property
    
Public Property Let idPcGrupo(ByVal value As Long)
    mProps.idPcGrupo = value
    mbooModificado = True
End Property

Public Property Get idPcGrupo() As Long
    idPcGrupo = mProps.idPcGrupo
End Property

Public Property Let idElemento(ByVal value As Long)
    mProps.idElemento = value
    mbooModificado = True
End Property

Public Property Get idElemento() As Long
    idElemento = mProps.idElemento
End Property

Public Property Let IdCuenta(ByVal value As Long)
    mProps.IdCuenta = value
    mbooModificado = True
End Property

Public Property Get IdCuenta() As Long
    IdCuenta = mProps.IdCuenta
End Property


Public Property Let FechaGrabado(ByVal value As Date)
    If Not (IsDate(value)) Then Err.Raise ERR_INVALIDO, "CuentasPCGrupo.FechaGrabado", MSGERR_INVALIDO
    mProps.FechaGrabado = value
    mbooModificado = True
End Property

Public Property Get FechaGrabado() As Date
    FechaGrabado = mProps.FechaGrabado
End Property

Public Property Get EsNuevo() As Boolean
    EsNuevo = mbooEsNuevo
End Property

Public Property Get Modificado() As Boolean
    Modificado = mbooModificado
End Property

Private Sub Class_Initialize()
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objPCG = Nothing
End Sub

'*** METODOS
Public Sub Grabar()
        Dim sql As String, rs As Recordset
        On Error GoTo Etiq
            
        If mbooEsNuevo Then
            sql = "SELECT * FROM CuentaPcGrupoPre WHERE 1=0"
            Set rs = Me.objPCG.Empresa.OpenRecordsetParaEdit(sql)
            rs.AddNew
            
        Else
            sql = "Select * FROM CuentaPCGrupoPre WHERE idpcgrupo = '" & Me.idPcGrupo & "'"
            Set rs = Me.objPCG.Empresa.OpenRecordsetParaEdit(sql)
        End If
        With rs
            rs!idPcGrupo = Me.idPcGrupo
            !idElemento = Me.idElemento
            !IdCuenta = Me.IdCuenta
            !FechaGrabado = Now
            .Update
            .Close
        End With
        Set rs = Nothing
        mbooEsNuevo = False
        mbooModificado = False
        Exit Sub
Etiq:
        Err.Raise Err.Number, Err.Source, Err.Description
        Exit Sub
End Sub




