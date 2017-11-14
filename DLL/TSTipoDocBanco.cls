VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TSTipoDocBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Empresa"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit



Private Type T_PROPS
    IdTipoDoc As Long
    CodTipoDoc As String
    Descripcion As String
    Comentario As String
    PermitePostFechado As Boolean
    UsoEnIngreso As Boolean
    BandCalculoBanco As Integer
    BandCalculoDoc As Integer
    BandValida As Boolean
    BandEliminable As Boolean
    DatosAdicionales As Boolean
    Efectivo As Boolean
    
    FechaGrabado As Date
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean


Public Sub Grabar()
    Dim sql As String, rs As Recordset
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código            '*** MAKOTO 17/oct/00
    sql = "SELECT CodTipoDoc FROM TSTipoDocBanco " & _
          "WHERE (CodTipoDoc='" & mProps.CodTipoDoc & "') AND " & _
                "(IdTipoDoc<>" & mProps.IdTipoDoc & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "TSTipoDoc", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM TSTipoDocBanco WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM TSTipoDocBanco WHERE IdTipoDoc=" & mProps.IdTipoDoc
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "TSTipoDocBanco_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en Pos
    With rs
        !CodTipoDoc = mProps.CodTipoDoc
        !Descripcion = mProps.Descripcion
        !Comentario = mProps.Comentario
        !PermitePostFechado = mProps.PermitePostFechado
        !UsoEnIngreso = mProps.UsoEnIngreso
        !BandCalculoBanco = mProps.BandCalculoBanco
        !BandCalculoDoc = mProps.BandCalculoDoc
        !BandValida = mProps.BandValida
        !BandEliminable = mProps.BandEliminable
        !DatosAdicionales = mProps.DatosAdicionales 'jeaa 17/07/2009
        !Efectivo = mProps.Efectivo
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.FechaGrabado = !FechaGrabado
        mProps.IdTipoDoc = !IdTipoDoc
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_DOCBAN", "Documento de Banco Nuevo: " & _
                                                         mProps.CodTipoDoc, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_DOCBAN", "Documento de Banco Modificada: " & _
                                                         mProps.CodTipoDoc, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_TIPODOCBANCO
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM TSTipoDocBanco WHERE CodTipoDoc='" & key & "'"
    Else
        sql = "SELECT * FROM TSTipoDocBanco WHERE IdTipoDoc=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IdTipoDoc = !IdTipoDoc
        mProps.CodTipoDoc = !CodTipoDoc
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!Comentario) Then mProps.Comentario = !Comentario
        If Not IsNull(!PermitePostFechado) Then mProps.PermitePostFechado = !PermitePostFechado
        If Not IsNull(!UsoEnIngreso) Then mProps.UsoEnIngreso = !UsoEnIngreso
        If Not IsNull(!BandCalculoBanco) Then mProps.BandCalculoBanco = !BandCalculoBanco
        If Not IsNull(!BandCalculoDoc) Then mProps.BandCalculoDoc = !BandCalculoDoc
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!BandEliminable) Then mProps.BandEliminable = !BandEliminable
        If Not IsNull(!DatosAdicionales) Then mProps.DatosAdicionales = !DatosAdicionales 'Agregado por jeaa 17/07/2009
        If Not IsNull(!Efectivo) Then mProps.Efectivo = !Efectivo
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset

    'Verifica si existe Kardex
    sql = "SELECT IdTipoDoc FROM TSKardex WHERE IdTipoDoc=" & mProps.IdTipoDoc
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "TSTipoDocBanco.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing
    
    'Si no es eliminable
    If Not mProps.BandEliminable Then
        Err.Raise ERR_NOELIMINABLE, "TSTipoDocBanco_Eliminar", MSGERR_NOELIMINABLE
    End If

    'Elimina la empresa
    sql = "DELETE FROM TSTipoDocBanco WHERE CodTipoDoc='" & mProps.CodTipoDoc & "'"
    mProps.objEmpresa.Execute sql, True

    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_DOCBAN", "Documento de Banco Eliminado: " & _
                                                         mProps.CodTipoDoc, "DT"

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_TIPODOCBANCO
End Sub

Public Function Clone() As TSTipoDocBanco
    Dim nuevo As TSTipoDocBanco
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New TSTipoDocBanco
    nuevo.Props = mProps
    nuevo.BandEliminable = True
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property




Public Property Get IdTipoDoc() As Long
    IdTipoDoc = mProps.IdTipoDoc
End Property



Public Property Let CodTipoDoc(ByVal value As String)
    If Len(value) > 5 Then
        Err.Raise ERR_INVALIDO, "TSTipoDocBanco_CodTipoDoc", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodTipoDoc = value
    mbooModificado = True
End Property

Public Property Get CodTipoDoc() As String
    CodTipoDoc = Trim$(mProps.CodTipoDoc)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "TSTipoDocBanco_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let Comentario(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "TSTipoDocBanco_Comentario", MSGERR_INVALIDO
    End If
    
    mProps.Comentario = value
    mbooModificado = True
End Property

Public Property Get Comentario() As String
    Comentario = Trim$(mProps.Comentario)
End Property

Public Property Let PermitePostFechado(ByVal value As Boolean)
    mProps.PermitePostFechado = value
    mbooModificado = True
End Property

Public Property Get PermitePostFechado() As Boolean
    PermitePostFechado = mProps.PermitePostFechado
End Property

Public Property Let UsoEnIngreso(ByVal value As Boolean)
    mProps.UsoEnIngreso = value
    mbooModificado = True
End Property

Public Property Get UsoEnIngreso() As Boolean
    UsoEnIngreso = mProps.UsoEnIngreso
End Property

Public Property Let BandCalculoBanco(ByVal value As Integer)
    If value < -1 Or value > 1 Then
        Err.Raise ERR_INVALIDO, "TSTipoDocBanco_BandCalculoBanco", MSGERR_INVALIDO
    End If
    
    mProps.BandCalculoBanco = value
    mbooModificado = True
End Property

Public Property Get BandCalculoBanco() As Integer
    BandCalculoBanco = mProps.BandCalculoBanco
End Property

Public Property Let BandCalculoDoc(ByVal value As Integer)
    If value < -1 Or value > 1 Then
        Err.Raise ERR_INVALIDO, "TSTipoDocBanco_BandCalculoDoc", MSGERR_INVALIDO
    End If
    
    mProps.BandCalculoDoc = value
    mbooModificado = True
End Property

Public Property Get BandCalculoDoc() As Integer
    BandCalculoDoc = mProps.BandCalculoDoc
End Property

Public Property Let BandValida(ByVal value As Boolean)
    mProps.BandValida = value
    mbooModificado = True
End Property

Public Property Get BandValida() As Boolean
    BandValida = mProps.BandValida
End Property

Public Property Let BandEliminable(ByVal value As Boolean)
    mProps.BandEliminable = value
    mbooModificado = True
End Property

Public Property Get BandEliminable() As Boolean
    BandEliminable = mProps.BandEliminable
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
    With mProps
        .BandCalculoBanco = 1
        .BandCalculoDoc = 1
        .BandEliminable = True
        .BandValida = True
        .PermitePostFechado = True
        .UsoEnIngreso = True
    End With

    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Property Let DatosAdicionales(ByVal value As Boolean)
    mProps.DatosAdicionales = value
    mbooModificado = True
End Property

Public Property Get DatosAdicionales() As Boolean
    DatosAdicionales = mProps.DatosAdicionales
End Property


'AUC para sacar un total de efectivo en los tipos de cobros
Public Property Let Efectivo(ByVal value As Boolean)
    mProps.Efectivo = value
    mbooModificado = True
End Property
Public Property Get Efectivo() As Boolean
    Efectivo = mProps.Efectivo
End Property

