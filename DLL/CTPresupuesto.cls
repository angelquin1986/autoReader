VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTPresupuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Empresa"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Private Type T_PROPS
    Anio As Integer
    Mes As Byte
    IdCuenta As Long
    CodCuenta As String
    Debe As Currency
    Haber As Currency
        
    FechaGrabado As Date
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean


Public Sub Grabar()
    Dim sql As String
    Dim rs As Recordset
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM CTPresupuesto WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM CTPresupuesto WHERE " & _
                 "Anio=" & mProps.Anio & " AND Mes=" & mProps.Mes & _
                 " AND IdCuetna=" & mProps.IdCuenta
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "CTPresupuesto_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en Pos
    With rs
        !Anio = mProps.Anio
        !Mes = mProps.Mes
        !IdCuenta = mProps.IdCuenta
        !Debe = mProps.Debe
        !Haber = mProps.Haber
        
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Friend Sub Recuperar()
    Dim sql As String, rs As Recordset
    
    sql = "SELECT A.*,B.CodCuenta FROM CTPresupuesto A " & _
          "INNER JOIN CTCuenta B ON A.IdCuenta=B.IdCuenta " & _
          "WHERE Anio=" & mProps.Anio & " AND Mes=" & mProps.Mes & _
             " AND A.IdCuetna=" & mProps.IdCuenta
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        If Not IsNull(!CodCuenta) Then mProps.CodCuenta = !CodCuenta
        If Not IsNull(!Debe) Then mProps.Debe = !Debe
        If Not IsNull(!Haber) Then mProps.Haber = !Haber
        
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String

    'Elimina la empresa
    sql = "DELETE FROM CTPresupuesto WHERE " & _
             "Anio=" & mProps.Anio & " AND Mes=" & mProps.Mes & _
             " AND IdCuetna=" & mProps.IdCuenta
    mProps.objEmpresa.Execute sql, True

    mbooEsNuevo = True
    mbooModificado = True
End Sub

Friend Function Clone() As CTPresupuesto
    Dim nuevo As CTPresupuesto
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New CTPresupuesto
    nuevo.Props = mProps
    
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






Public Property Let Anio(ByVal value As Integer)
    If Not mbooEsNuevo Then
        Err.Raise ERR_NOMODIFICABLE, "CTPresupuesto_Anio", MSGERR_NOMODIFICABLE
    End If
    
    mProps.Anio = value
    mbooModificado = True
End Property

Public Property Get Anio() As Integer
    Anio = mProps.Anio
End Property

Public Property Let Mes(ByVal value As Byte)
    If Not mbooEsNuevo Then
        Err.Raise ERR_NOMODIFICABLE, "CTPresupuesto_Mes", MSGERR_NOMODIFICABLE
    End If
    
    If value < 0 Or value > 12 Then
        Err.Raise ERR_INVALIDO, "CTPresupuesto_Mes", MSGERR_INVALIDO
    End If
    
    mProps.Mes = value
    mbooModificado = True
End Property

Public Property Get Mes() As Byte
    Mes = mProps.Mes
End Property

Public Property Let IdCuenta(ByVal value As Long)
    mProps.IdCuenta = value
    mbooModificado = True
End Property

Public Property Get IdCuenta() As Long
    IdCuenta = mProps.IdCuenta
End Property

Public Property Let CodCuenta(ByVal value As String)
    Dim sql As String
    Dim rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCuenta Then
        'Actualiza IdCuenta también
        sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "CTPresupuesto_CodCuenta", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            IdCuenta = rs!IdCuenta
        End If
        rs.Close
    End If

    Set rs = Nothing
    mProps.CodCuenta = value
    mbooModificado = True
End Property

Public Property Get CodCuenta() As String
    CodCuenta = Trim$(mProps.CodCuenta)
End Property



Public Property Let Debe(ByVal value As Currency)
    mProps.Debe = value
    mbooModificado = True
End Property

Public Property Get Debe() As Currency
    Debe = mProps.Debe
End Property

Public Property Let Haber(ByVal value As Currency)
    mProps.Haber = value
    mbooModificado = True
End Property

Public Property Get Haber() As Currency
    Haber = mProps.Haber
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
    Set mProps.objEmpresa = Nothing
End Sub




