VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FichaMedEnfDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdExmFisDetalle As Long
    IdExmFisico As Long
    idEnfermedad As Long
    codEnfermedad As String
    Descripcion As String
    Orden As Integer
    objFichaMedExamen As FichaMedicaExamenes
    CodigoEnfermedad As String
    DescripcionEnfermedad As String
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean
Event Refresh(Tipo As Integer)

' /* Cargar Enfermedades */
Public Property Let CodigoEnfermedad(ByVal value As String)
    mProps.CodigoEnfermedad = value
    mbooModificado = True
End Property

Public Property Get CodigoEnfermedad() As String
    CodigoEnfermedad = mProps.CodigoEnfermedad
End Property

Public Property Let DescripcionEnfermedad(ByVal value As String)
    mProps.DescripcionEnfermedad = value
    mbooModificado = True
End Property

Public Property Get DescripcionEnfermedad() As String
    DescripcionEnfermedad = mProps.DescripcionEnfermedad
End Property
' /* Cargar Enfermedades */

Public Property Get IdExmFisDetalle() As Long
    IdExmFisDetalle = mProps.IdExmFisDetalle
End Property

Public Property Let IdExmFisico(ByVal value As Long)
    mProps.IdExmFisico = value
    mbooModificado = True
End Property

Public Property Get IdExmFisico() As Long
    IdExmFisico = mProps.IdExmFisico
End Property

Public Property Let idEnfermedad(ByVal value As Long)
    mProps.idEnfermedad = value
    mbooModificado = True
End Property

Public Property Get idEnfermedad() As Long
    idEnfermedad = mProps.idEnfermedad
End Property

Friend Property Set objFichaMedExamen(ByVal obj As FichaMedicaExamenes)
    Set mProps.objFichaMedExamen = obj
End Property

Public Property Get objFichaMedExamen() As FichaMedicaExamenes
    Set objFichaMedExamen = mProps.objFichaMedExamen
End Property
Public Property Let codEnfermedad(ByVal value As String)
     Dim sql As String
    Dim rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.codEnfermedad Then
        If Len(value) > 0 Then
            'Actualiza IdCentro también
            'sql = "SELECT IdEnfermedad FROM FichaEnfermedad WHERE CodEnfermedad='" & value & "'"
            sql = "SELECT IdCuenta FROM FichaEnfermedad WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objFichaMedExamen.mobjEmpleado.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Ficha.CodCanton", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                'idEnfermedad = rs!idEnfermedad
                idEnfermedad = rs!IdCuenta
            End If
            rs.Close
            Set rs = Nothing
        Else
            idEnfermedad = 0
        End If
        mProps.codEnfermedad = value
        mbooModificado = True
    End If
End Property

Public Property Get codEnfermedad() As String
    codEnfermedad = mProps.codEnfermedad
End Property

Public Property Let Descripcion(ByVal value As String)
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = mProps.Descripcion
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
    Set mProps.objFichaMedExamen = Nothing
End Sub

Public Sub Recuperar(ByVal rs As Recordset)
    Dim sql As String
   
    With rs
            mProps.IdExmFisDetalle = !IdExmFisDetalle
            mProps.IdExmFisico = !IdExmFisico
            mProps.idEnfermedad = !idEnfermedad

            If Not IsNull(!Orden) Then mProps.Orden = !Orden
    End With
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Grabar()
    Dim sql As String, rs As Recordset
    On Error GoTo Etiq
    
        sql = "SELECT * FROM FichaMedEnfDetalle WHERE IdExmFisDetalle = " & mProps.IdExmFisDetalle
        Set rs = mProps.objFichaMedExamen.mobjEmpleado.Empresa.OpenRecordsetParaEdit(sql)
    
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sql = "SELECT * FROM dbo.FichaMedEnfDetalle WHERE 1=0"
        Set rs = mProps.objFichaMedExamen.mobjEmpleado.Empresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    Else
    End If
    
    With rs
            !IdExmFisico = mProps.objFichaMedExamen.IdExmFisico
            !idEnfermedad = mProps.idEnfermedad
            !Orden = mProps.Orden
            
            .Update
            .Move 0

            mProps.IdExmFisDetalle = !IdExmFisDetalle
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


Public Function EliminaDetalle(cod, idficha) As fichaDetalle
    Dim nuevo As fichaDetalle
    Set nuevo = New fichaDetalle
    Set nuevo.objFicha = Me
    nuevo.BorradoFisico Val(cod), Val(idficha)
    Set EliminaDetalle = nuevo
    Set nuevo = Nothing
End Function


'Public Function Eliminar() As Boolean
'    Dim sql As String
'    Dim rs As Recordset
'    On Error GoTo CapturaError
'    sql = "DELETE FROM dbo.FichaMedicaExmFisico WHERE idEmpleado = " & Me.IdEmpleado
'
'
'    mbooEsNuevo = True
'    mbooModificado = True
'    Exit Function
'
'CapturaError:
'    MsgBox Err.Description
'    Set rs = Nothing
'    Eliminar = False
'    Exit Function
'End Function

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property

Public Property Let Orden(ByVal value As Integer)
    mProps.Orden = value
    mbooModificado = True
End Property

