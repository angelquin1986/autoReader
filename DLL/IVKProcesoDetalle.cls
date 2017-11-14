VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVKProcesoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROP

    idpd As Long
    Transid As Long
    idproceso As Long
    CodProceso As String
    Nota As String
    Ruta As String
    objEmpresa As Empresa
End Type
Private mProps As T_PROP
Public Sub grabar()
    Dim sql As String, rs As Recordset
    sql = "SELECT * FROM ivkProcesoDetalle WHERE 1=0"
    Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !Transid = mProps.Transid
        !idproceso = mProps.idproceso
        !Nota = mProps.Nota
        !Ruta = mProps.Ruta
        .Update
        
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.idpd = !idpd             'Guarda el Id(Autonumerico)
        .Close
    End With
    Set rs = Nothing
'    mProps.objEmpresa.RaiseRefreshEvent REFRESH_DETALLEPROCESO
End Sub

Friend Sub recuperar(rs As Recordset)
    With rs
        If Not .EOF Then
            mProps.idpd = !idpd
            mProps.Transid = !Transid
            mProps.idproceso = !idproceso
            If Not IsNull(!Nota) Then mProps.Nota = !Nota
            If Not IsNull(!Ruta) Then mProps.Ruta = !Ruta
        End If
    End With
End Sub

Friend Function Clone() As IVKProcesoDetalle
    Dim nuevo As IVKProcesoDetalle
    Set nuevo = New IVKProcesoDetalle
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(Value As T_PROP)          'Para el método Clone
    mProps = Value
    mProps.idpd = 0
End Property

Friend Property Set objEmpresa(Value As Empresa)
    Set mProps.objEmpresa = Value
End Property
Public Property Get objEmpresa() As Empresa
    Set objEmpresa = mProps.objEmpresa
End Property
'AUC cambiado para el reproceso de recetas 19/10/06
Public Property Get idpd() As Long
    idpd = mProps.idpd
End Property
Public Property Get Transid() As Long
    Transid = mProps.Transid
End Property

Public Property Let Transid(ByVal Value As Long)
    mProps.Transid = Value
End Property

Private Sub Class_Initialize()
  Set objEmpresa = gobjMain.EmpresaActual
End Sub
Public Property Get Ruta() As String
    Ruta = mProps.Ruta
End Property
Public Property Let Ruta(ByVal Value As String)
    If Len(Value) > 120 Then
        Err.Raise ERR_INVALIDO, "ivkprocesodetalle.ruta", MSGERR_INVALIDO
    End If
    mProps.Ruta = Value
    'mbooModificado = True
End Property
Public Property Get Nota() As String
    Nota = mProps.Nota
End Property
Public Property Let Nota(ByVal Value As String)
    If Len(Value) > 120 Then
        Err.Raise ERR_INVALIDO, "ivkprocesodetalle.Nota", MSGERR_INVALIDO
    End If
    mProps.Nota = Value
'    mbooModificado = True
End Property

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    'Elimina el Contrato de costo
    sql = "DELETE FROM ivkProcesoDetalle  WHERE Idpd=" & mProps.idpd
    mProps.objEmpresa.Execute sql, True
    
    mProps.objEmpresa.GrabaGNLogAccion "ELI_ivkProcesoDetalle", "Proceso Detalle Eliminado ", "DT"
   
    'mProps.objEmpresa.RaiseRefreshEvent REFRESH_PROTIEMPOS
End Sub

Public Property Get idproceso() As Long
    idproceso = mProps.idproceso
End Property

Public Property Let idproceso(ByVal Value As Long)
    mProps.idproceso = Value
End Property

Public Property Let CodProceso(ByVal Value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If Value <> mProps.CodProceso Then
        If Len(Value) > 0 Then
            
            sql = "SELECT IdProceso FROM ivproceso WHERE Codproceso='" & Value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "ivkprocesodetalle_CodProceso", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                idproceso = rs!idproceso
            End If
            rs.Close
        Else
            idproceso = 0
        End If
        
        mProps.CodProceso = Value
        'mbooModificado = True
    End If
    Set rs = Nothing
End Property
