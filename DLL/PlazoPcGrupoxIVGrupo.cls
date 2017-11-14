VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlazoPcGrupoxIVGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***jeaa  30/12/2004
Private Type T_PROPS
    IdGrupoDetalle As Long
    idPcGrupo  As Integer
    IDIVGrupo  As Integer
    Valor As Currency
    CodPCGrupo  As String
    CodIVGrupo  As String
    FechaGrabado As Date
    ValorAdicional As Currency
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Public Sub Grabar()
    Dim sql As String, rs As Recordset
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    'Verifica si no repite el código
        
        sql = "SELECT Valor FROM PlazoIVGPCG " & _
          "WHERE (CodPCGrupo='" & mProps.CodPCGrupo & "') AND " & _
          " (codIVGrupo='" & mProps.CodIVGrupo & "') AND " & _
                "(IdGrupoDetalle<>" & mProps.IdGrupoDetalle & ")"

    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "Plazo PCGrupo x IVGrupo", MSGERR_REPITECODIGO
    End If
    rs.Close
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM PlazoIVGPCG WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM PlazoIVGPCG WHERE IdGrupoDetalle=" & mProps.IdGrupoDetalle
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)

        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "DesctPCGxIVG_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro
    With rs
        !Valor = mProps.Valor
        !idPcGrupo = mProps.idPcGrupo
        !CodPCGrupo = mProps.CodPCGrupo
        !IDIVGrupo = mProps.IDIVGrupo
        !CodIVGrupo = mProps.CodIVGrupo
        !ValorAdicional = mProps.ValorAdicional
        !FechaGrabado = Now
        .Update
        .Move 0             'Para actualizar
        mProps.FechaGrabado = !FechaGrabado
        mProps.IdGrupoDetalle = !IdGrupoDetalle
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_DPLAZOPCGXIVG
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String
    Dim rs As Recordset
    Dim v As Variant
    Dim NumIVGrupo As Integer, NumPCGrupo As String
    NumIVGrupo = CInt(mProps.objEmpresa.GNOpcion.ObtenerValor("PlazoIVGrupo")) + 1
    NumPCGrupo = mProps.objEmpresa.GNOpcion.ObtenerValor("PlazoPCGrupo")

    v = Split(key, ",")
        If VarType(key) = vbString Then
            sql = "SELECT "
            sql = sql & " desct.IdGrupoDetalle , desct.IDPCGrupo, pcg.Coddiascredito as  CodPCGrupo"
            sql = sql & " , desct.IDIVGrupo, ivg.CodGrupo" & NumIVGrupo & " as  CodIVGrupo, desct.Valor , desct.FechaGrabado, desct.ValorAdicional "
            sql = sql & " FROM pcdiascredito pcg"
            sql = sql & " inner join PlazoIVGPCG desct"
            sql = sql & " inner join ivgrupo" & NumIVGrupo & " ivg"
            sql = sql & " on desct.idivgrupo=ivg.idgrupo" & NumIVGrupo
            sql = sql & " on desct.idpcgrupo=pcg.iddiascredito"
            sql = sql & " WHERE desct.codPCGrupo='" & v(0) & "'and desct.codIVGrupo='" & v(1) & "'"
        
        End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'Recupera los datos de Usuario
    With rs
        mProps.IdGrupoDetalle = !IdGrupoDetalle
        mProps.idPcGrupo = !idPcGrupo
        mProps.IDIVGrupo = !IDIVGrupo
        mProps.CodPCGrupo = !CodPCGrupo
        mProps.CodIVGrupo = !CodIVGrupo
        mProps.Valor = !Valor
        If Not IsNull(!ValorAdicional) Then mProps.ValorAdicional = !ValorAdicional
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
  
    'Elimina el registro
    sql = "DELETE FROM PlazoIVGPCG WHERE IdGrupoDetalle=" & mProps.IdGrupoDetalle
    mProps.objEmpresa.Execute sql, True

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_DPCGXIVG
End Sub

Public Function Clone() As PlazoPcGrupoxIVGrupo
    Dim nuevo As PlazoPcGrupoxIVGrupo
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New PlazoPcGrupoxIVGrupo
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdGrupoDetalle = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Get IdGrupoDetalle() As Long
    IdGrupoDetalle = mProps.IdGrupoDetalle
End Property


Public Property Let idPcGrupo(ByVal value As Integer)
    mProps.idPcGrupo = value
    mbooModificado = True
End Property

Public Property Get idPcGrupo() As Integer
    idPcGrupo = mProps.idPcGrupo
End Property

Public Property Let IDIVGrupo(ByVal value As Integer)
    mProps.IDIVGrupo = value
    mbooModificado = True
End Property

Public Property Get IDIVGrupo() As Integer
    IDIVGrupo = mProps.IDIVGrupo
End Property

Public Property Let CodPCGrupo(ByVal value As String)
    Dim sql As String, rs As Recordset
    Dim NumPCGrupo As String
    NumPCGrupo = mProps.objEmpresa.GNOpcion.ObtenerValor("PlazoNumPCGrupo")
    
    'Cuando cambia el valor
    If value <> mProps.CodPCGrupo Then
        sql = "SELECT IdDiasCredito FROM pcdiascredito WHERE CoddiasCredito='" & value & "'"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "CodPcGRupo", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            idPcGrupo = rs.Fields(0)
        End If
        rs.Close
    End If
    
    Set rs = Nothing


    mProps.CodPCGrupo = value
    mbooModificado = True
End Property

Public Property Get CodPCGrupo() As String
    CodPCGrupo = mProps.CodPCGrupo
End Property

Public Property Let CodIVGrupo(ByVal value As String)
    Dim sql As String, rs As Recordset
    Dim NumIVGrupo As Integer
    NumIVGrupo = CInt(mProps.objEmpresa.GNOpcion.ObtenerValor("PlazoIVGrupo")) + 1
    
    'Cuando cambia el valor
    If value <> mProps.CodIVGrupo Then
        sql = "SELECT IdGrupo" & NumIVGrupo & " FROM ivgrupo" & NumIVGrupo & " WHERE CodGrupo" & NumIVGrupo & "='" & value & "'"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "CodivGRupo", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            IDIVGrupo = rs.Fields(0)
        End If
        rs.Close
    End If
    
    Set rs = Nothing


    mProps.CodIVGrupo = value
    mbooModificado = True
End Property

Public Property Get CodIVGrupo() As String
    CodIVGrupo = mProps.CodIVGrupo
End Property


Public Property Let Valor(ByVal value As Currency)
    mProps.Valor = value
    mbooModificado = True
End Property

Public Property Get Valor() As Currency
    Valor = mProps.Valor
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

Friend Property Let Modificado(ByVal value As Boolean)
    mbooModificado = value
End Property


Private Sub Class_Initialize()
    'Valores predeterminados
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Property Let ValorAdicional(ByVal value As Currency)
    mProps.ValorAdicional = value
    mbooModificado = True
End Property

Public Property Get ValorAdicional() As Currency
    ValorAdicional = mProps.ValorAdicional
End Property



