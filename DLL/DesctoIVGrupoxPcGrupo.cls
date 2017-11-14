VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DesctoIVGrupoxPcGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***jeaa  30/12/2004
Private Type T_PROPS
    IdGrupoDetalle As Long
    IDPCGrupo  As Integer
    Valor As Currency
    IDIVGrupo  As Integer
    Orden As Long
    CodPCGrupo As String      'revisar  si nos sirve
    CodIVGrupo As String      'revisar  si nos sirve
    NumPCGrupo  As Byte
    NumIVGrupo  As Byte
    FechaGrabado As Date
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mcolDesctoIVxPCDetalle As Collection  '***jeaa 30/12/2004

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

'Establece relacion  con el PADRE
Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property


Public Function Clone() As DesctoIVGrupoxPcGrupo
    Dim nuevo As DesctoIVGrupoxPcGrupo
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New DesctoIVGrupoxPcGrupo
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mbooModificado = True
End Property



Public Property Get Orden() As Long
    Orden = mProps.Orden
End Property

Public Property Let Orden(ByVal value As Long)
    mProps.Orden = value
End Property

Public Property Let Valor(ByVal value As Currency)
    mProps.Valor = value
mbooModificado = True
End Property

Public Property Get Valor() As Currency
    Valor = mProps.Valor
End Property


Public Property Get IDPCGrupo() As Integer
    IDPCGrupo = mProps.IDPCGrupo
End Property

Public Property Let IDPCGrupo(ByVal value As Integer)
    mProps.IDPCGrupo = value
    mbooModificado = True
End Property
Public Property Get IDIVGrupo() As Integer
    IDIVGrupo = mProps.IDIVGrupo
End Property

Public Property Let IDIVGrupo(ByVal value As Integer)
    mProps.IDIVGrupo = value
  mbooModificado = True
End Property


Private Sub Class_Initialize()
    'Valores predeterminados
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Sub Recuperar()
    Dim sql As String, rs As Recordset, i As Long, mod_antes As Boolean
    Dim nuevo As DesctoIVGrupoxPcGrupo
    mod_antes = mbooModificado      'Guarda el estado de Modificado
    If mcolDesctoIVxPCDetalle Is Nothing Then Set mcolDesctoIVxPCDetalle = New Collection
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("DescuentoNumPCGrupo")) > 0 Then
        mProps.NumPCGrupo = CInt((gobjMain.EmpresaActual.GNOpcion.ObtenerValor("DescuentoNumPCGrupo"))) + 1
    End If
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("DescuentoNumIVGrupo")) > 0 Then
        mProps.NumIVGrupo = CInt((gobjMain.EmpresaActual.GNOpcion.ObtenerValor("DescuentoNumIVGrupo"))) + 1
    End If
    BorrarDetalle
    
    'Recupera los
    sql = "SELECT * FROM DesctoIVGrupoPCGrupo ORDER BY IdGrupoDetalle"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
                Set nuevo = New DesctoIVGrupoxPcGrupo
                mcolDesctoIVxPCDetalle.Add nuevo
                i = mcolDesctoIVxPCDetalle.Count
'                Set DesctoIVGrupoxPcGrupo(i).Empresa = Me
                DesctoIVGrupoxPcGrupo(i).RecuperarDEtalle rs
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    mbooModificado = mod_antes      'Para que no cambie el estado de Modificado
End Sub

Public Property Get DesctoIVGrupoxPcGrupo(ix As Long) As DesctoIVGrupoxPcGrupo
    If mcolDesctoIVxPCDetalle Is Nothing Then Recuperar   'posilbe recursividad
    Set DesctoIVGrupoxPcGrupo = mcolDesctoIVxPCDetalle.Item(ix)
End Property

Public Property Get NumRecargoDetalle() As Long
    If mcolDesctoIVxPCDetalle Is Nothing Then
        NumRecargoDetalle = 0
    Else
        NumRecargoDetalle = mcolDesctoIVxPCDetalle.Count
    End If
End Property


'--------------- Manejo de coleccion de RecargoDetalle
Public Function AddDetalleDescto() As Long
    Dim nuevo As DesctoIVGrupoxPcGrupo
    
    If mcolDesctoIVxPCDetalle Is Nothing Then Recuperar
    Set nuevo = New DesctoIVGrupoxPcGrupo
    mcolDesctoIVxPCDetalle.Add nuevo
    AddDetalleDescto = mcolDesctoIVxPCDetalle.Count   'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function

Public Sub Grabar()
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As DesctoIVGrupoxPcGrupo, sql As String
    'Elimina  todos  los registros
    sql = "Delete DesctoIVGrupoPCGrupo  "
    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    If Not mcolDesctoIVxPCDetalle Is Nothing Then
        sql = "SELECT * FROM DesctoIVGrupoPCGrupo WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        For i = 1 To mcolDesctoIVxPCDetalle.Count
            rs1.AddNew
            Set obj = mcolDesctoIVxPCDetalle.Item(i)
            rs1!Valor = obj.Valor
            rs1!IDPCGrupo = obj.IDPCGrupo
            rs1!IDIVGrupo = obj.IDIVGrupo
'            rs1!Orden = obj.Orden
            rs1!FechaGrabado = Now
            rs1.Update
        Next i
    End If
    mbooEsNuevo = False
    mbooModificado = False
    Set rs1 = Nothing
    Set obj = Nothing
End Sub

Public Sub RemoveRecargoDetalle(ByVal ix As Long, Optional obj As DesctoIVGrupoxPcGrupo)
    If mcolDesctoIVxPCDetalle Is Nothing Then Recuperar
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolDesctoIVxPCDetalle.Count
            If mcolDesctoIVxPCDetalle.Item(ix) Is obj Then
                mcolDesctoIVxPCDetalle.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
          mcolDesctoIVxPCDetalle.Remove ix
    End If
    
    'mcolDetalleFamilia.Remove ix
    mbooModificado = True
End Sub




Private Sub BorrarDetalle()
    'Borra  toda  la informacion de la coleccion
     If mcolDesctoIVxPCDetalle Is Nothing Then Set mcolDesctoIVxPCDetalle = New Collection
    With mcolDesctoIVxPCDetalle
        Do While .Count > 0
            .Remove .Count
        Loop
    End With
    mbooModificado = True
End Sub

Public Property Let CodPCGrupo(value As String)
    Dim rs As Recordset, sql As String
    mProps.CodPCGrupo = value
    If value = "" Then
        mProps.IDPCGrupo = 0
    Else
        'actualiza  el IDGrupo
        sql = "Select IdGrupo" & mProps.NumPCGrupo & " From PCGrupo" & mProps.NumPCGrupo & _
              " Where CodGrupo" & mProps.NumPCGrupo & " = '" & value & "'"
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
        If Not rs.EOF Then
            mProps.IDPCGrupo = rs.Fields(0).value
        End If
        rs.Close
        Set rs = Nothing
    End If
'    mProps.objIVRecargo.Modificado = True
End Property

Public Property Get CodPCGrupo() As String
    CodPCGrupo = mProps.CodPCGrupo
End Property

Public Property Let NumPCGrupo(ByVal value As Byte)
    mProps.NumPCGrupo = value
'    mProps.Modificado = True
End Property

Public Property Get NumPCGrupo() As Byte
    NumPCGrupo = mProps.NumPCGrupo
End Property


Public Property Let NumIVGrupo(ByVal value As Byte)
    mProps.NumIVGrupo = value
'    mProps.Modificado = True
End Property

Public Property Get NumIVGrupo() As Byte
    NumIVGrupo = mProps.NumIVGrupo
End Property

Public Property Let CodIVGrupo(value As String)
    Dim rs As Recordset, sql As String
    mProps.CodIVGrupo = value
    If value = "" Then
        mProps.IDPCGrupo = 0
    Else
        'actualiza  el IDGrupo
        sql = "Select IdGrupo" & mProps.NumIVGrupo & " From IVGrupo" & mProps.NumIVGrupo & _
              " Where CodGrupo" & mProps.NumIVGrupo & " = '" & value & "'"
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
'        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then
            mProps.IDIVGrupo = rs.Fields(0).value
        End If
        rs.Close
        Set rs = Nothing
    End If
'    mProps.objIVRecargo.Modificado = True
End Property

Public Property Get CodIVGrupo() As String
    CodIVGrupo = mProps.CodIVGrupo
End Property

Friend Sub RecuperarDEtalle(rs As Recordset)
    Dim sql As String, rs1  As Recordset
        If Len(mProps.NumPCGrupo) = 0 Then
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("DescuentoNumPCGrupo")) > 0 Then
                mProps.NumPCGrupo = CInt((gobjMain.EmpresaActual.GNOpcion.ObtenerValor("DescuentoNumPCGrupo")))
            End If
        Else
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("DescuentoNumPCGrupo")) > 0 Then
                mProps.NumPCGrupo = CInt((gobjMain.EmpresaActual.GNOpcion.ObtenerValor("DescuentoNumPCGrupo"))) + 1
            End If
        End If
        If Len(mProps.NumIVGrupo) = 0 Then
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("DescuentoNumIVGrupo")) > 0 Then
                mProps.NumIVGrupo = CInt((gobjMain.EmpresaActual.GNOpcion.ObtenerValor("DescuentoNumIVGrupo")))
            End If
        Else
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("DescuentoNumIVGrupo")) > 0 Then
                mProps.NumIVGrupo = CInt((gobjMain.EmpresaActual.GNOpcion.ObtenerValor("DescuentoNumIVGrupo"))) + 1
            End If

      End If
    With rs
        If Not .EOF Then
            mProps.IdGrupoDetalle = !IdGrupoDetalle
            If Not IsNull(!IDPCGrupo) Then mProps.IDPCGrupo = !IDPCGrupo
            If Not IsNull(!IDIVGrupo) Then mProps.IDIVGrupo = !IDIVGrupo
            mProps.Valor = !Valor
            mProps.FechaGrabado = !FechaGrabado
'            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            'recupera  el codigo  de PCGrupo
            On Error Resume Next
                sql = "Select CodGrupo" & mProps.NumPCGrupo & " From PCGrupo" & mProps.NumPCGrupo & _
                      " Where IDGrupo" & mProps.NumPCGrupo & " = " & mProps.IDPCGrupo
                Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
                If Not rs1.EOF Then
                    mProps.CodPCGrupo = rs1.Fields(0).value
                End If
                    'recupera  el codigo  de IVGrupo
                sql = "Select CodGrupo" & mProps.NumIVGrupo & " From IVGrupo" & mProps.NumIVGrupo & _
                      " Where IDGrupo" & mProps.NumIVGrupo & " = " & mProps.IDIVGrupo
                Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
                If Not rs1.EOF Then
                    mProps.CodIVGrupo = rs1.Fields(0).value
                End If
                rs1.Close
                Set rs1 = Nothing
        End If
    End With
End Sub

Public Property Get EsNuevo() As Boolean
    EsNuevo = mbooEsNuevo
End Property


Public Property Get Modificado() As Boolean
    Modificado = mbooModificado
End Property

Friend Property Let Modificado(ByVal value As Boolean)
    mbooModificado = value
End Property

Public Property Get FechaGrabado() As Date
    FechaGrabado = mProps.FechaGrabado
End Property
