VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GNEstadoProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_Props
    idEstado As Long
    CodEstado As String
    Descripcion As String
    bandValida As Boolean
    Menu As String
    Valor As Integer
    color As String
    BandTerminado As Boolean
    FechaGrabado As Date
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mProps As T_Props
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean


Public Sub grabar()
    Dim sql As String, rs As Recordset
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código                '*** MAKOTO 17/oct/00
    sql = "SELECT CodEstado FROM GNEstadoProd " & _
          "WHERE (Codestado='" & mProps.CodEstado & "') AND " & _
                "(Idestado<>" & mProps.idEstado & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "GNEstadoProd", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM GNEstadoProd WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM GNEstadoProd WHERE IDEstado=" & mProps.idEstado
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "GNEstadoProd.Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro
    With rs
        !CodEstado = mProps.CodEstado
        !Descripcion = mProps.Descripcion
        !Menu = mProps.Menu
        !Valor = mProps.Valor
        !bandValida = mProps.bandValida
        !BandTerminado = mProps.BandTerminado
        !color = mProps.color
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.FechaGrabado = !FechaGrabado
        mProps.idEstado = !idEstado
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_CTLOC", "Estado Nuevo: " & _
                                                         mProps.CodEstado, "CT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_CTLOC", "Estado Modificado: " & _
                                                         mProps.CodEstado, "CT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_GNESTADOPROD
End Sub

Friend Sub recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM GNEstadoProd  WHERE CodEstado='" & key & "'"
    Else
        sql = "SELECT * FROM GNEstadoProd WHERE Idestado=" & key
    End If

    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    rs.MoveFirst   'para que se coloque al inicio jeaa
    'Recupera los datos
    With rs
        mProps.idEstado = !idEstado
        mProps.CodEstado = !CodEstado
        mProps.Valor = !Valor
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!Menu) Then mProps.Menu = !Menu
        If Not IsNull(!bandValida) Then mProps.bandValida = !bandValida
        If Not IsNull(!BandTerminado) Then mProps.BandTerminado = !BandTerminado
        If Not IsNull(!color) Then mProps.color = !color
         
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    
    'Verifica si ya existe una transacción usando el mismo
    sql = "SELECT ivkp.transid,gne.codestado as estado,gne1.codestado as estado1 "
    sql = sql & "FROM ivkproceso ivkp  "
    sql = sql & "Inner join gnestadoprod gne on gne.valor =  ivkp.estado  "
    sql = sql & "Inner join gnestadoprod gne1 on gne1.valor =  ivkp.estado1   "
    sql = sql & " Where gne.idEstado = " & mProps.idEstado
    sql = sql & " or gne1.Idestado=" & mProps.idEstado
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "GNEstadoProd.Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close

    'Elimina el registro
    sql = "DELETE FROM GNEstadoProd WHERE Idestado=" & mProps.idEstado
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_GNEstadoProd", "Estado Eliminado: " & _
                                                         mProps.CodEstado, "DT"

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_GNESTADOPROD
End Sub

Friend Function Clone() As GNEstadoProd
    Dim nuevo As GNEstadoProd

    'Copia las propiedades al nuevo objeto
    Set nuevo = New GNEstadoProd
    nuevo.Props = mProps

    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(Value As T_Props)         'Para el método Clone
    mProps = Value
    mProps.idEstado = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Get idEstado() As Long
    idEstado = mProps.idEstado
End Property


Public Property Let CodEstado(ByVal Value As String)
    If Len(Value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNGasto_Codestado", MSGERR_INVALIDO
    End If
    ValidaCodigo Value      'Valida si contiene caracteres no validos
    
    mProps.CodEstado = Value
    mbooModificado = True
End Property

Public Property Get CodEstado() As String
    CodEstado = Trim$(mProps.CodEstado)
End Property

Public Property Let Descripcion(ByVal Value As String)
    If Len(Value) > 80 Then
        Err.Raise ERR_INVALIDO, "CodGasto_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = Value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let bandValida(ByVal Value As Boolean)
    mProps.bandValida = Value
    mbooModificado = True
End Property

Public Property Get bandValida() As Boolean
    bandValida = mProps.bandValida
End Property
Public Property Let BandTerminado(ByVal Value As Boolean)
    mProps.BandTerminado = Value
    mbooModificado = True
End Property

Public Property Get BandTerminado() As Boolean
    BandTerminado = mProps.BandTerminado
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
    'Valores iniciales
    mProps.bandValida = True
    mbooEsNuevo = True
    mbooModificado = False
    
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Property Let Menu(ByVal Value As String)
    If Len(Value) > 20 Then
        Err.Raise ERR_INVALIDO, "CodGasto_menu", MSGERR_INVALIDO
    End If
    
    mProps.Menu = Value
    mbooModificado = True
End Property

Public Property Get Menu() As String
    Menu = Trim$(mProps.Menu)
End Property


Public Property Let Valor(ByVal Value As Integer)
    mProps.Valor = Value
    mbooModificado = True
End Property

Public Property Get Valor() As Integer
    Valor = Trim$(mProps.Valor)
End Property


Public Property Let color(ByVal Value As String)
    mProps.color = Value
    mbooModificado = True
End Property

Public Property Get color() As String
    color = Trim$(mProps.color)
End Property



Friend Sub RecuperarxMenu(key As Variant)
    Dim sql As String, rs As Recordset
    On Error GoTo CapturaError
    sql = "SELECT * FROM GNEstadoProd  WHERE menu='" & key & "'"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    rs.MoveFirst   'para que se coloque al inicio jeaa
    'Recupera los datos
    With rs
        mProps.idEstado = !idEstado
        mProps.CodEstado = !CodEstado
        mProps.Valor = !Valor
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!Menu) Then mProps.Menu = !Menu
        If Not IsNull(!bandValida) Then mProps.bandValida = !bandValida
        If Not IsNull(!BandTerminado) Then mProps.BandTerminado = !BandTerminado
        If Not IsNull(!color) Then mProps.color = !color
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    Exit Sub
CapturaError:
    rs.Close
    MsgBox Err.Description
    Set rs = Nothing
End Sub

Friend Sub RecuperarxColor(key As Variant)
    Dim sql As String, rs As Recordset
    On Error GoTo CapturaError
    sql = "SELECT * FROM GNEstadoProd  WHERE color='" & key & "'"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    rs.MoveFirst   'para que se coloque al inicio jeaa
    'Recupera los datos
    With rs
        mProps.idEstado = !idEstado
        mProps.CodEstado = !CodEstado
        mProps.Valor = !Valor
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!Menu) Then mProps.Menu = !Menu
        If Not IsNull(!bandValida) Then mProps.bandValida = !bandValida
        If Not IsNull(!BandTerminado) Then mProps.BandTerminado = !BandTerminado
        If Not IsNull(!color) Then mProps.color = !color
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    Exit Sub
CapturaError:
    rs.Close
    MsgBox Err.Description
    Set rs = Nothing
End Sub

