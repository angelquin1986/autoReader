VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type T_PROPS
    PosID As String
    Descripcion As String
    FechaGrabado As Date
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Public Sub Grabar()
    Dim sql As String, i As Long
    Dim rs As Recordset
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
#If DAOLIB = 0 Then
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
#End If

    'Si es nueva
    If mbooEsNuevo Then
        'Verifica si no repite el código
        sql = "SELECT PosID FROM Pos WHERE PosID='" & mProps.PosID & "'"
#If DAOLIB Then
        Set rs = gobjMain.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
        rs.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly
#End If
        If Not rs.EOF Then      'Si ya existe el codigo
            Err.Raise ERR_REPITECODIGO, "Pos", MSGERR_REPITECODIGO
        End If
        rs.Close
        
        sql = "SELECT * FROM Pos WHERE 1=0"
#If DAOLIB Then
        Set rs = gobjMain.Database.OpenRecordset(sql)
#Else
        rs.Open sql, gobjMain.Coneccion, adOpenKeyset, adLockPessimistic
#End If
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM Pos WHERE PosID='" & mProps.PosID & "'"
#If DAOLIB Then
        Set rs = gobjMain.Database.OpenRecordset(sql)
        rs.Edit
#Else
        rs.Open sql, gobjMain.Coneccion, adOpenKeyset, adLockPessimistic
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "Pos_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en Pos
    With rs
        !PosID = PosID
        !Descripcion = Descripcion
        
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Friend Sub Recuperar()
    Dim sql As String, rs As Recordset
    
    sql = "SELECT * FROM Pos WHERE PosID='" & mProps.PosID & "'"
    
#If DAOLIB Then
    Set rs = gobjMain.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly
#End If
    'Recupera los datos
    With rs
        If Not IsNull(!PosID) Then mProps.PosID = !PosID
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String

    'Si no tiene derecho de supervisor, no puede hacer
    If Not gobjUsuarioActual.BandSupervisor Then
        Err.Raise ERR_NODERECHO, "Pos_Eliminar", MSGERR_NODERECHO
    End If
    
    'Elimina la caja
    sql = "DELETE FROM Pos WHERE PosID='" & mProps.PosID & "'"
#If DAOLIB Then
    gobjMain.Database.Execute sql, dbFailOnError
#Else
    gobjMain.Coneccion.Execute sql
#End If

    mbooEsNuevo = True
    mbooModificado = True
End Sub

Public Function Clone() As Pos
    Dim nuevo As Pos
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New Pos
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mbooModificado = True
End Property






Public Property Let PosID(ByVal value As String)
    If Not mbooEsNuevo Then
        Err.Raise ERR_NOMODIFICABLE, "Pos_PosID", MSGERR_NOMODIFICABLE
    End If
    
    If Len(value) > 5 Then
        Err.Raise ERR_INVALIDO, "Pos_PosID", MSGERR_INVALIDO
    End If
    
    mProps.PosID = value
    mbooModificado = True
End Property

Public Property Get PosID() As String
    PosID = Trim$(mProps.PosID)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "Pos_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
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


