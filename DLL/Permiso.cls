VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Permiso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Operador"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

 Option Explicit

Private Type T_PROP
    Id As Long
    objGrupo As Grupo               'Referencia al objeto padre
    CodEmpresa As String            'Codigo de empresa

    CatCentroCostoVer As Boolean
    CatPlanCuentaVer As Boolean
    CatBancoVer As Boolean
    CatInventarioVer As Boolean
    CatBodegaVer As Boolean
    CatVendedorVer As Boolean
    CatProveedorVer As Boolean
    CatClienteVer As Boolean
    CatAFVer As Boolean
    CatResponsableVer As Boolean
    CatInfEmpresaVer As Boolean
    CatRolVer As Boolean
    CatInventarioCostoVer As Boolean

    CatCentroCostoMod As Boolean
    CatPlanCuentaMod As Boolean
    CatBancoMod As Boolean
    CatInventarioMod As Boolean
    CatBodegaMod As Boolean
    CatVendedorMod As Boolean
    CatProveedorMod As Boolean
    CatClienteMod As Boolean
    CatAFMod As Boolean
    CatResponsableMod As Boolean
    CatInfEmpresaMod As Boolean
    CatRolMod As Boolean
    CatInventarioPrecioMod As Boolean

    ConsAsiento As Boolean
    ConsCentroCosto As Boolean
    ConsBanco As Boolean
    ConsInventario As Boolean
    ConsTrans As Boolean
    ConsVenta As Boolean
    ConsProveedor As Boolean
    ConsCliente As Boolean
    ConsAF As Boolean
    ConsRol As Boolean
    ConsConciliacion As Boolean
    VerCostos As Boolean        '*********** jeaa 30-03-04
    CatUnidadVer As Boolean 'AUC Para unidad 12/12/05
    CatUnidadMod As Boolean 'AUC  12/12/05
    'Caption de menus
    MenuCaption(1 To 9) As String
    MenuReporteCaption(1 To 9) As String
    ConsRepVen As Boolean 'AUC Mostrar Reporte por Vendedor/usuario

End Type

Private mProps As T_PROP
Private mcolTrans As Collection
Private mcolMenu As Collection
Private mcolMenuReporte As Collection



Friend Sub Grabar()
    Dim cat1 As Long, cat2 As Long, cons As Long, i As Long, sql As String
    Dim rs As Recordset
    
    mProps.objGrupo.GenerarProcesando "Está grabándo permiso para '" & mProps.CodEmpresa & "'"
    
    sql = "SELECT * FROM Permiso WHERE 1=0"
#If DAOLIB Then
    Set rs = gobjMain.Database.OpenRecordset(sql, , dbAppendOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.Open sql, gobjMain.Coneccion, adOpenKeyset, adLockPessimistic
#End If
    rs.AddNew
    rs!CodGrupo = mProps.objGrupo.CodGrupo
    rs!CodEmpresa = mProps.CodEmpresa
    
    With rs
        'Permisos de ver catalogos
        cat1 = mProps.CatCentroCostoVer * -&H1
        cat1 = cat1 + mProps.CatPlanCuentaVer * -&H2
        cat1 = cat1 + mProps.CatBancoVer * -&H4
        cat1 = cat1 + mProps.CatInventarioVer * -&H8
        cat1 = cat1 + mProps.CatBodegaVer * -&H10
        cat1 = cat1 + mProps.CatVendedorVer * -&H20
        cat1 = cat1 + mProps.CatProveedorVer * -&H40
        cat1 = cat1 + mProps.CatClienteVer * -&H80
        cat1 = cat1 + mProps.CatAFVer * -&H100
        cat1 = cat1 + mProps.CatResponsableVer * -&H200
        cat1 = cat1 + mProps.CatInfEmpresaVer * -&H400
        cat1 = cat1 + mProps.CatRolVer * -&H800
        cat1 = cat1 + mProps.CatInventarioCostoVer * -&H1000
        
        'Permisos de Modificar catalogos
        cat2 = mProps.CatCentroCostoMod * -&H1
        cat2 = cat2 + mProps.CatPlanCuentaMod * -&H2
        cat2 = cat2 + mProps.CatBancoMod * -&H4
        cat2 = cat2 + mProps.CatInventarioMod * -&H8
        cat2 = cat2 + mProps.CatBodegaMod * -&H10
        cat2 = cat2 + mProps.CatVendedorMod * -&H20
        cat2 = cat2 + mProps.CatProveedorMod * -&H40
        cat2 = cat2 + mProps.CatClienteMod * -&H80
        cat2 = cat2 + mProps.CatAFMod * -&H100
        cat2 = cat2 + mProps.CatResponsableMod * -&H200
        cat2 = cat2 + mProps.CatInfEmpresaMod * -&H400
        cat2 = cat2 + mProps.CatRolMod * -&H800
        cat2 = cat2 + mProps.CatInventarioPrecioMod * -&H1000
        
        'Permisos de ver consultas
        cons = mProps.ConsAsiento * -&H1
        cons = cons + mProps.ConsCentroCosto * -&H2
        cons = cons + mProps.ConsBanco * -&H4
        cons = cons + mProps.ConsInventario * -&H8
        cons = cons + mProps.ConsTrans * -&H10
        cons = cons + mProps.ConsVenta * -&H20
        cons = cons + mProps.ConsProveedor * -&H40
        cons = cons + mProps.ConsCliente * -&H80
        cons = cons + mProps.ConsAF * -&H100
        cons = cons + mProps.ConsRol * -&H200
        cons = cons + mProps.ConsConciliacion * -&H400
        cons = cons + mProps.ConsRepVen * -&H800 'AUC 20/07/2010
    
        
        .Fields("PermisoVerCatalogos") = cat1
        .Fields("PermisoModCatalogos") = cat2
        .Fields("PermisoConsultas") = cons
        
        'Caption de menus
        For i = 1 To 9
            .Fields("MenuCaption" & i) = mProps.MenuCaption(i)
        Next i
        
        'Caption de menus Reporte
        For i = 1 To 9
            .Fields("MenuReporteCaption" & i) = mProps.MenuReporteCaption(i)
        Next i
'        .Fields("MenuReporteCaption" & 1) = mProps.MenuReporteCaption(1)
        
        
        .Update
#If DAOLIB Then
        .Bookmark = .LastModified
#End If
        .Move 0
        mProps.Id = !Id
        .Close
    End With
    
    'Graba todos los PermisoTrans
    For i = 1 To CountTrans
        mProps.objGrupo.GenerarProcesando "Está grabándo permisos para '" & mProps.CodEmpresa & "'. " & i & " de " & CountTrans     '*** MAKOTO 27/dic/00
        Trans(i).Grabar
    Next i
    
    'Graba todos los PermisoMenu
    For i = 1 To CountMenu
        mProps.objGrupo.GenerarProcesando "Está grabándo menús para '" & mProps.CodEmpresa & "'. " & i & " de " & CountMenu      '*** MAKOTO 27/dic/00
        Menu(i).Grabar
    Next i
    
    'Graba todos los PermisoMenuReporte
    For i = 1 To CountMenuReporte
        mProps.objGrupo.GenerarProcesando "Está grabándo menús reporte para '" & mProps.CodEmpresa & "'. " & i & " de " & CountMenuReporte      '*** MAKOTO 27/dic/00
        MenuReporte(i).Grabar
    Next i
    
    Set rs = Nothing
End Sub

Friend Sub Recuperar()
    Dim sql As String, rs As Recordset
  
    sql = "SELECT * FROM Permiso " & _
          "WHERE CodGrupo='" & mProps.objGrupo.CodGrupo & "'" & _
             " AND CodEmpresa='" & mProps.CodEmpresa & "'"
#If DAOLIB Then
    Set rs = gobjMain.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly
#End If
    RecuperarRS rs
    Set rs = Nothing
End Sub


Friend Sub RecuperarRS(ByVal rs As Recordset)
    Dim sql As String, i As Long
    Dim cat1 As Long, cat2 As Long, cons As Long
    Dim rs2 As Recordset
  
    With rs
        mProps.Id = !Id
        If Not IsNull(!PermisoVerCatalogos) Then cat1 = !PermisoVerCatalogos
        If Not IsNull(!PermisoModCatalogos) Then cat2 = !PermisoModCatalogos
        If Not IsNull(!PermisoConsultas) Then cons = !PermisoConsultas
        
        'Permisos de ver catalogos
        mProps.CatCentroCostoVer = CBool(cat1 And &H80000001)
        mProps.CatPlanCuentaVer = CBool(cat1 And &H80000002)
        mProps.CatBancoVer = CBool(cat1 And &H80000004)
        mProps.CatInventarioVer = CBool(cat1 And &H80000008)
        mProps.CatBodegaVer = CBool(cat1 And &H80000010)
        mProps.CatVendedorVer = CBool(cat1 And &H80000020)
        mProps.CatProveedorVer = CBool(cat1 And &H80000040)
        mProps.CatClienteVer = CBool(cat1 And &H80000080)
        mProps.CatAFVer = CBool(cat1 And &H80000100)
        mProps.CatResponsableVer = CBool(cat1 And &H80000200)
        mProps.CatInfEmpresaVer = CBool(cat1 And &H80000400)
        mProps.CatRolVer = CBool(cat1 And &H80000800)
        mProps.CatInventarioCostoVer = CBool(cat1 And &H80001000)
        
        
        'Permisos de Modificar catalogos
        mProps.CatCentroCostoMod = CBool(cat2 And &H80000001)
        mProps.CatPlanCuentaMod = CBool(cat2 And &H80000002)
        mProps.CatBancoMod = CBool(cat2 And &H80000004)
        mProps.CatInventarioMod = CBool(cat2 And &H80000008)
        mProps.CatBodegaMod = CBool(cat2 And &H80000010)
        mProps.CatVendedorMod = CBool(cat2 And &H80000020)
        mProps.CatProveedorMod = CBool(cat2 And &H80000040)
        mProps.CatClienteMod = CBool(cat2 And &H80000080)
        mProps.CatAFMod = CBool(cat2 And &H80000100)
        mProps.CatResponsableMod = CBool(cat2 And &H80000200)
        mProps.CatInfEmpresaMod = CBool(cat2 And &H80000400)
        mProps.CatRolMod = CBool(cat2 And &H80000800)
        mProps.CatInventarioPrecioMod = CBool(cat2 And &H80001000)
        
        'Permisos de ver consultas
        mProps.ConsAsiento = CBool(cons And &H80000001)
        mProps.ConsCentroCosto = CBool(cons And &H80000002)
        mProps.ConsBanco = CBool(cons And &H80000004)
        mProps.ConsInventario = CBool(cons And &H80000008)
        mProps.ConsTrans = CBool(cons And &H80000010)
        mProps.ConsVenta = CBool(cons And &H80000020)
        mProps.ConsProveedor = CBool(cons And &H80000040)
        mProps.ConsCliente = CBool(cons And &H80000080)
        mProps.ConsAF = CBool(cons And &H80000100)
        mProps.ConsRol = CBool(cons And &H80000200)
        mProps.ConsConciliacion = CBool(cons And &H80000400)
        mProps.ConsRepVen = CBool(cons And &H80000800)
        
        'Caption de menus
        For i = 1 To 9
            If Not IsNull(.Fields("MenuCaption" & i)) Then
                mProps.MenuCaption(i) = Trim$(.Fields("MenuCaption" & i))
            Else
                mProps.MenuCaption(i) = ""
            End If
        Next i
    
        'Caption de menusReporte
        For i = 1 To 9
            If Not IsNull(.Fields("MenuReporteCaption" & i)) Then
                mProps.MenuReporteCaption(i) = Trim$(.Fields("MenuReporteCaption" & i))
            Else
                mProps.MenuReporteCaption(i) = ""
            End If
        Next i
    
    End With
    
    'Recupera los PermisoTrans
    LimpiaColeccion mcolTrans
    'jeaa 20/08/2005 Agrgado order by codtrans
    sql = "SELECT * FROM PermisoTrans WHERE IdPermiso=" & mProps.Id & " order by codtrans"
#If DAOLIB Then
    Set rs2 = gobjMain.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs2 = New ADODB.Recordset
    rs2.CursorLocation = adUseClient
    rs2.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly
#End If
    With rs2
        Do Until .EOF
            i = AddTrans
            Trans(i).CodTrans = !CodTrans
            Trans(i).RecuperarRS rs2
            .MoveNext
        Loop
        .Close
    End With
    
    'Recupera los PermisoMenu
    LimpiaColeccion mcolMenu
    sql = "SELECT * FROM PermisoMenu WHERE IdPermiso=" & mProps.Id & " ORDER BY NumMenu,Orden"
#If DAOLIB Then
    Set rs2 = gobjMain.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    rs2.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly
#End If
    With rs2
        Do Until .EOF
            i = AddMenu
            Menu(i).RecuperarRS rs2
            .MoveNext
        Loop
        .Close
    End With
'    Set rs2 = Nothing

    'Recupera los PermisoMenuReporte
        LimpiaColeccion mcolMenuReporte
    sql = "SELECT * FROM PermisoMenuReporte WHERE IdPermiso=" & mProps.Id & " ORDER BY NumMenu,Orden"
#If DAOLIB Then
    Set rs2 = gobjMain.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    rs2.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly
#End If
    With rs2
        Do Until .EOF
            i = AddMenuReporte
            MenuReporte(i).RecuperarRS rs2
            .MoveNext
        Loop
        .Close
    End With
    Set rs2 = Nothing


End Sub


Public Function Clone() As Permiso
    Dim nuevo As Permiso, i As Long, nt As PermisoTrans, nm As PermisoMenu, nmr As PermisoMenuReporte
    
    Set nuevo = New Permiso
    nuevo.Props = mProps
    
    'Copia los PermisoTrans al nuevo objeto
    For i = 1 To mcolTrans.Count
        Set nt = Trans(i).Clone
        nuevo.AddTransParaClone nt
    Next i
    
    'Copia los PermisoMenu al nuevo objeto
    For i = 1 To mcolMenu.Count
        Set nm = Menu(i).Clone
        nuevo.AddMenuParaClone nm
    Next i
    
    'Copia los PermisoMenuReporte al nuevo objeto
    For i = 1 To mcolMenuReporte.Count
        Set nmr = MenuReporte(i).Clone
        nuevo.AddMenuReporteParaClone nmr
    Next i
    
    Set Clone = nuevo
    Set nuevo = Nothing
    Set nt = Nothing
    Set nm = Nothing
    Set nmr = Nothing
End Function


Friend Sub AddTransParaClone(obj As PermisoTrans)
    Set obj.Permiso = Me
    mcolTrans.Add obj
End Sub

Friend Sub AddMenuParaClone(obj As PermisoMenu)
    Set obj.Permiso = Me
    mcolMenu.Add obj
End Sub

Friend Property Let Props(value As T_PROP)          'Para el método Clone
    mProps = value
    mProps.objGrupo.SetModificado
End Property

Friend Property Set Grupo(value As Grupo)     'Para el método AddPermiso de Grupo
    Set mProps.objGrupo = value
End Property

Public Property Get Grupo() As Grupo
    Set Grupo = mProps.objGrupo
End Property

'========= Manejo de PermisoTrans
Public Function AddTrans() As Long
    Dim nuevo As PermisoTrans
    
    Set nuevo = New PermisoTrans
    Set nuevo.Permiso = Me
    mcolTrans.Add nuevo
    AddTrans = mcolTrans.Count     'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mProps.objGrupo.SetModificado
End Function

Public Sub RemoveTrans(ix As Long)
    mcolTrans.Remove ix
    mProps.objGrupo.SetModificado
End Sub

'Duplica un Trans para el mismo Grupo
Public Function CopiaTrans(ix As Long) As Long
    Dim nuevo As PermisoTrans, obj As PermisoTrans
    
    Set obj = mcolTrans.item(ix)
    Set nuevo = obj.Clone
    mcolTrans.Add nuevo
    
    CopiaTrans = mcolTrans.Count       'Devuelve el indice del nuevo elemento
    Set nuevo = Nothing
    Set obj = Nothing
    mProps.objGrupo.SetModificado
End Function

Public Property Get CountTrans() As Long
    CountTrans = mcolTrans.Count
End Property

Public Property Get Trans(ByVal key As Variant) As PermisoTrans
    Dim obj As PermisoTrans, i As Long

    If VarType(key) = vbObject Then
        Set obj = key
        i = BuscaTrans(obj.CodTrans)
    ElseIf VarType(key) = vbString Then
        i = BuscaTrans(CStr(key))
    Else
        i = key
    End If
    If i > 0 Then
        Set Trans = mcolTrans.item(i)
    Else
        Set Trans = Nothing
    End If
End Property

Public Function BuscaTrans(CodT As String) As Long
    Dim i As Long, obj As PermisoTrans
    
    BuscaTrans = 0
    For i = 1 To mcolTrans.Count
        Set obj = mcolTrans.item(i)
        If obj.CodTrans = CodT Then
            BuscaTrans = i
            Exit For
        End If
    Next i
End Function

Public Sub LimpiaTrans()
    Dim i As Long
    
    For i = mcolTrans.Count To 1 Step -1
        mcolTrans.Remove i
    Next i
End Sub



Public Function ListaTrans(ByVal cadena As Boolean, Optional ByVal Modulo As String) As Variant
    Dim i As Long, s As String, v() As Variant, cnt As Integer
    Dim pt As PermisoTrans, sql As String, rs As Recordset
    
    If cadena Then
        If mcolTrans.Count = 1 Then
            s = mcolTrans.item(1).CodTrans
        ElseIf mcolTrans.Count > 0 Then
            For i = 1 To mcolTrans.Count
                Set pt = mcolTrans.item(i)
                If Len(s) > 0 Then s = s & ", "
                s = s & "'" & pt.CodTrans & "'"
            Next i
            s = "IN (" & s & ")"
        End If
        ListaTrans = s
    Else
        
        For i = 1 To mcolTrans.Count
            Set pt = mcolTrans.item(i)
            sql = "SELECT NombreTrans FROM GNTrans WHERE CodTrans='" & pt.CodTrans & "'"
            If Len(Modulo) > 0 Then sql = sql & " AND Modulo='" & Modulo & "'"
            
#If DAOLIB Then
            Set rs = gobjMain.EmpresaActual.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.Open sql, gobjMain.EmpresaActual.Coneccion, adOpenStatic, adLockReadOnly
#End If
            If Not rs.EOF Then
                ReDim Preserve v(1, cnt)
                v(0, cnt) = pt.CodTrans
                v(1, cnt) = rs!NombreTrans
                cnt = cnt + 1
            End If
            rs.Close
        Next i
        
        ListaTrans = v
        Set pt = Nothing
        Set rs = Nothing
    End If
End Function



'=========== Manejo de PermisoMenu
Public Function AddMenu() As Long
    Dim nuevo As PermisoMenu
    
    Set nuevo = New PermisoMenu
    Set nuevo.Permiso = Me
    mcolMenu.Add nuevo
    AddMenu = mcolMenu.Count     'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mProps.objGrupo.SetModificado
End Function

Public Sub RemoveMenu(ix As Variant)
    Dim obj As PermisoMenu, i As Long

    'Si llama con numero de indice del objeto
    If (VarType(ix) = vbLong) Or (VarType(ix) = vbInteger) Then
        mcolMenu.Remove ix      'Elimina el objeto
        
    'Si llama con la referencia al objeto que quiere eliminar
    Else
        'Busca el objeto dentro de la colección
        For i = mcolMenu.Count To 1 Step -1
            Set obj = mcolMenu.item(i)
            If obj Is ix Then       'Si encuentra
                mcolMenu.Remove i       'Elimina el objeto
                GoTo eliminado
            End If
        Next i
        'Si no encuentra el objeto, genera error
        Err.Raise ERR_INVALIDO, "Permiso_RemoveMenu", "No se puede eliminar menu."
    End If
eliminado:
    mProps.objGrupo.SetModificado
    Set obj = Nothing
End Sub

'Duplica un Menu para el mismo Grupo
Public Function CopiaMenu(ix As Long) As Long
    Dim nuevo As PermisoMenu, obj As PermisoMenu
    
    Set obj = mcolMenu.item(ix)
    Set nuevo = obj.Clone
    mcolMenu.Add nuevo
    
    CopiaMenu = mcolMenu.Count       'Devuelve el indice del nuevo elemento
    Set nuevo = Nothing
    Set obj = Nothing
    mProps.objGrupo.SetModificado
End Function

Public Property Get CountMenu() As Long
    CountMenu = mcolMenu.Count
End Property

Public Property Get Menu(ix As Long) As PermisoMenu
    Set Menu = mcolMenu.item(ix)
End Property

Public Sub LimpiaMenu()
    Dim i As Long
    
    For i = mcolMenu.Count To 1 Step -1
        mcolMenu.Remove i
    Next i
End Sub








Public Sub SetModificado()
    mProps.objGrupo.SetModificado
End Sub







Public Property Get Id() As Long
    Id = mProps.Id
End Property

Public Property Let CodEmpresa(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "Permiso_CodEmpresa", MSGERR_INVALIDO
    End If
    
    mProps.CodEmpresa = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CodEmpresa() As String
    CodEmpresa = Trim$(mProps.CodEmpresa)
End Property

Public Property Let CatCentroCostoVer(ByVal value As Boolean)
    mProps.CatCentroCostoVer = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatCentroCostoVer() As Boolean
    CatCentroCostoVer = mProps.CatCentroCostoVer
End Property

Public Property Let CatPlanCuentaVer(ByVal value As Boolean)
    mProps.CatPlanCuentaVer = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatPlanCuentaVer() As Boolean
    CatPlanCuentaVer = mProps.CatPlanCuentaVer
End Property

Public Property Let CatBancoVer(ByVal value As Boolean)
    mProps.CatBancoVer = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatBancoVer() As Boolean
    CatBancoVer = mProps.CatBancoVer
End Property

Public Property Let CatInventarioVer(ByVal value As Boolean)
    mProps.CatInventarioVer = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatInventarioVer() As Boolean
    CatInventarioVer = mProps.CatInventarioVer
End Property

Public Property Let CatBodegaVer(ByVal value As Boolean)
    mProps.CatBodegaVer = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatBodegaVer() As Boolean
    CatBodegaVer = mProps.CatBodegaVer
End Property

Public Property Let CatVendedorVer(ByVal value As Boolean)
    mProps.CatVendedorVer = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatVendedorVer() As Boolean
    CatVendedorVer = mProps.CatVendedorVer
End Property

Public Property Let CatProveedorVer(ByVal value As Boolean)
    mProps.CatProveedorVer = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatProveedorVer() As Boolean
    CatProveedorVer = mProps.CatProveedorVer
End Property

Public Property Let CatClienteVer(ByVal value As Boolean)
    mProps.CatClienteVer = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatClienteVer() As Boolean
    CatClienteVer = mProps.CatClienteVer
End Property

Public Property Let CatAFVer(ByVal value As Boolean)
    mProps.CatAFVer = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatAFVer() As Boolean
    CatAFVer = mProps.CatAFVer
End Property

Public Property Let CatResponsableVer(ByVal value As Boolean)
    mProps.CatResponsableVer = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatResponsableVer() As Boolean
    CatResponsableVer = mProps.CatResponsableVer
End Property

Public Property Let CatInfEmpresaVer(ByVal value As Boolean)
    mProps.CatInfEmpresaVer = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatInfEmpresaVer() As Boolean
    CatInfEmpresaVer = mProps.CatInfEmpresaVer
End Property

Public Property Let CatRolVer(ByVal value As Boolean)
    mProps.CatRolVer = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatRolVer() As Boolean
    CatRolVer = mProps.CatRolVer
End Property


Public Property Let CatInventarioCostoVer(ByVal value As Boolean)
    mProps.CatInventarioCostoVer = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatInventarioCostoVer() As Boolean
    CatInventarioCostoVer = mProps.CatInventarioCostoVer
End Property




Public Property Let CatCentroCostoMod(ByVal value As Boolean)
    mProps.CatCentroCostoMod = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatCentroCostoMod() As Boolean
    CatCentroCostoMod = mProps.CatCentroCostoMod
End Property

Public Property Let CatPlanCuentaMod(ByVal value As Boolean)
    mProps.CatPlanCuentaMod = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatPlanCuentaMod() As Boolean
    CatPlanCuentaMod = mProps.CatPlanCuentaMod
End Property

Public Property Let CatBancoMod(ByVal value As Boolean)
    mProps.CatBancoMod = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatBancoMod() As Boolean
    CatBancoMod = mProps.CatBancoMod
End Property

Public Property Let CatInventarioMod(ByVal value As Boolean)
    mProps.CatInventarioMod = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatInventarioMod() As Boolean
    CatInventarioMod = mProps.CatInventarioMod
End Property

Public Property Let CatBodegaMod(ByVal value As Boolean)
    mProps.CatBodegaMod = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatBodegaMod() As Boolean
    CatBodegaMod = mProps.CatBodegaMod
End Property

Public Property Let CatVendedorMod(ByVal value As Boolean)
    mProps.CatVendedorMod = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatVendedorMod() As Boolean
    CatVendedorMod = mProps.CatVendedorMod
End Property

Public Property Let CatProveedorMod(ByVal value As Boolean)
    mProps.CatProveedorMod = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatProveedorMod() As Boolean
    CatProveedorMod = mProps.CatProveedorMod
End Property

Public Property Let CatClienteMod(ByVal value As Boolean)
    mProps.CatClienteMod = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatClienteMod() As Boolean
    CatClienteMod = mProps.CatClienteMod
End Property

Public Property Let CatAFMod(ByVal value As Boolean)
    mProps.CatAFMod = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatAFMod() As Boolean
    CatAFMod = mProps.CatAFMod
End Property

Public Property Let CatResponsableMod(ByVal value As Boolean)
    mProps.CatResponsableMod = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatResponsableMod() As Boolean
    CatResponsableMod = mProps.CatResponsableMod
End Property

Public Property Let CatInfEmpresaMod(ByVal value As Boolean)
    mProps.CatInfEmpresaMod = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatInfEmpresaMod() As Boolean
    CatInfEmpresaMod = mProps.CatInfEmpresaMod
End Property

Public Property Let CatRolMod(ByVal value As Boolean)
    mProps.CatRolMod = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatRolMod() As Boolean
    CatRolMod = mProps.CatRolMod
End Property

Public Property Let CatInventarioPrecioMod(ByVal value As Boolean)
    mProps.CatInventarioPrecioMod = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatInventarioPrecioMod() As Boolean
    CatInventarioPrecioMod = mProps.CatInventarioPrecioMod
End Property




Public Property Let ConsAsiento(ByVal value As Boolean)
    mProps.ConsAsiento = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get ConsAsiento() As Boolean
    ConsAsiento = mProps.ConsAsiento
End Property

Public Property Let ConsCentroCosto(ByVal value As Boolean)
    mProps.ConsCentroCosto = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get ConsCentroCosto() As Boolean
    ConsCentroCosto = mProps.ConsCentroCosto
End Property

Public Property Let ConsBanco(ByVal value As Boolean)
    mProps.ConsBanco = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get ConsBanco() As Boolean
    ConsBanco = mProps.ConsBanco
End Property

Public Property Let ConsInventario(ByVal value As Boolean)
    mProps.ConsInventario = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get ConsInventario() As Boolean
    ConsInventario = mProps.ConsInventario
End Property

Public Property Let ConsTrans(ByVal value As Boolean)
    mProps.ConsTrans = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get ConsTrans() As Boolean
    ConsTrans = mProps.ConsTrans
End Property

Public Property Let ConsVenta(ByVal value As Boolean)
    mProps.ConsVenta = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get ConsVenta() As Boolean
    ConsVenta = mProps.ConsVenta
End Property

Public Property Let ConsProveedor(ByVal value As Boolean)
    mProps.ConsProveedor = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get ConsProveedor() As Boolean
    ConsProveedor = mProps.ConsProveedor
End Property

Public Property Let ConsCliente(ByVal value As Boolean)
    mProps.ConsCliente = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get ConsCliente() As Boolean
    ConsCliente = mProps.ConsCliente
End Property

Public Property Let ConsAF(ByVal value As Boolean)
    mProps.ConsAF = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get ConsAF() As Boolean
    ConsAF = mProps.ConsAF
End Property

Public Property Let ConsRol(ByVal value As Boolean)
    mProps.ConsRol = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get ConsRol() As Boolean
    ConsRol = mProps.ConsRol
End Property

Public Property Let ConsConciliacion(ByVal value As Boolean)
    mProps.ConsConciliacion = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get ConsConciliacion() As Boolean
    ConsConciliacion = mProps.ConsConciliacion
End Property

Public Property Let MenuCaption(index As Integer, ByVal value As String)
    If Len(value) > 20 Or index < 1 Or index > 9 Then
        Err.Raise ERR_INVALIDO, "Permiso_MenuCaption", MSGERR_INVALIDO
    End If
    
    mProps.MenuCaption(index) = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get MenuCaption(index As Integer) As String
    If index < 1 Or index > 9 Then
        Err.Raise ERR_INVALIDO, "Permiso_MenuCaption", MSGERR_INVALIDO
    End If
    MenuCaption = Trim$(mProps.MenuCaption(index))
End Property



Public Sub DarTodo()
    With mProps
        'Permisos de contabilidad
        .CatCentroCostoVer = True
        .CatPlanCuentaVer = True
        .CatBancoVer = True
        .CatInventarioVer = True
        .CatBodegaVer = True
        .CatVendedorVer = True
        .CatProveedorVer = True
        .CatClienteVer = True
        .CatAFVer = True
        .CatResponsableVer = True
        .CatInfEmpresaVer = True
        .CatRolVer = True
        .CatInventarioCostoVer = True
        
        .CatCentroCostoMod = True
        .CatPlanCuentaMod = True
        .CatBancoMod = True
        .CatInventarioMod = True
        .CatBodegaMod = True
        .CatVendedorMod = True
        .CatProveedorMod = True
        .CatClienteMod = True
        .CatAFMod = True
        .CatResponsableMod = True
        .CatInfEmpresaMod = True
        .CatRolMod = True
        .CatInventarioPrecioMod = True
    
        .ConsAsiento = True
        .ConsCentroCosto = True
        .ConsBanco = True
        .ConsInventario = True
        .ConsTrans = True
        .ConsVenta = True
        .ConsProveedor = True
        .ConsCliente = True
        .ConsAF = True
        .ConsRol = True
        .ConsConciliacion = True
    End With
End Sub


Private Sub Class_Initialize()
    Set mcolTrans = New Collection
    Set mcolMenu = New Collection
    Set mcolMenuReporte = New Collection
    DarTodo
End Sub

Private Sub Class_Terminate()
    Set mProps.objGrupo = Nothing
    Set mcolTrans = Nothing
    Set mcolMenu = Nothing
    Set mcolMenuReporte = Nothing
End Sub

Public Property Let VerCostos(ByVal value As Boolean)
    mProps.VerCostos = value
    mProps.objGrupo.SetModificado
End Property
Public Property Get VerCostos() As Boolean
    VerCostos = mProps.VerCostos
End Property

'AUC
Public Property Get CatUnidadVer() As Boolean
    CatUnidadVer = mProps.CatUnidadVer
End Property

Public Property Let CatUnidadVer(ByVal value As Boolean)
    mProps.CatUnidadVer = value
    mProps.objGrupo.SetModificado
End Property

'AUC 2/12/05
Public Property Let CatUnidadMod(ByVal value As Boolean)
    mProps.CatUnidadMod = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get CatUnidadMod() As Boolean
    CatUnidadMod = mProps.CatUnidadMod
End Property

Public Function ListaTransxTipoTrans(ByVal cadena As Boolean, Optional ByVal Modulo As String, Optional ByVal CodTrans As String) As Variant
End Function

Public Function ListaTransxTipoTransNew(ByVal cadena As Boolean, Optional ByVal Modulo As String, Optional ByVal CodTrans As String, Optional ByVal CompTrans As String) As Variant
    Dim i As Long, s As String, v() As Variant, cnt As Integer
    Dim pt As PermisoTrans, sql As String, rs As Recordset
    
    If cadena Then
        If mcolTrans.Count = 1 Then
            s = mcolTrans.item(1).CodTrans
        ElseIf mcolTrans.Count > 0 Then
            For i = 1 To mcolTrans.Count
                Set pt = mcolTrans.item(i)
                If Len(s) > 0 Then s = s & ", "
                s = s & "'" & pt.CodTrans & "'"
            Next i
            s = "IN (" & s & ")"
        End If
        ListaTransxTipoTransNew = s
    Else
        
        For i = 1 To mcolTrans.Count
            Set pt = mcolTrans.item(i)
            sql = "SELECT Descripcion FROM GNTrans WHERE CodTrans='" & pt.CodTrans & "'"
            If Len(Modulo) > 0 Then sql = sql & " AND Modulo='" & Modulo & "'"
            If Len(CodTrans) > 0 Then sql = sql & " AND  AnexoCodTipoTrans = '" & CodTrans & "' "
            If Len(CompTrans) > 0 Then sql = sql & " AND  AnexoCodTipoComp = '" & CompTrans & "' "
            
    
#If DAOLIB Then
            Set rs = gobjMain.EmpresaActual.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.Open sql, gobjMain.EmpresaActual.Coneccion, adOpenStatic, adLockReadOnly
#End If
            If Not rs.EOF Then
                ReDim Preserve v(1, cnt)
                v(0, cnt) = pt.CodTrans
                v(1, cnt) = rs!Descripcion
                cnt = cnt + 1
            End If
            rs.Close
        Next i
        
        ListaTransxTipoTransNew = v
        Set pt = Nothing
        Set rs = Nothing
    End If
End Function


Public Property Let MenuReporteCaption(index As Integer, ByVal value As String)
    If Len(value) > 20 Or index < 1 Or index > 9 Then
        Err.Raise ERR_INVALIDO, "Permiso_MenuReporteCaption", MSGERR_INVALIDO
    End If
    
    mProps.MenuReporteCaption(index) = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get MenuReporteCaption(index As Integer) As String
    If index < 1 Or index > 9 Then
        Err.Raise ERR_INVALIDO, "Permiso_MenuReporteCaption", MSGERR_INVALIDO
    End If
    MenuReporteCaption = Trim$(mProps.MenuReporteCaption(index))
End Property

'=========== Manejo de PermisoMenuReporte
Public Function AddMenuReporte() As Long
    Dim nuevo As PermisoMenuReporte
    
    Set nuevo = New PermisoMenuReporte
    Set nuevo.Permiso = Me
    mcolMenuReporte.Add nuevo
    AddMenuReporte = mcolMenuReporte.Count     'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mProps.objGrupo.SetModificado
End Function

Public Sub RemoveMenuReporte(ix As Variant)
    Dim obj As PermisoMenuReporte, i As Long

    'Si llama con numero de indice del objeto
    If (VarType(ix) = vbLong) Or (VarType(ix) = vbInteger) Then
        mcolMenuReporte.Remove ix      'Elimina el objeto
        
    'Si llama con la referencia al objeto que quiere eliminar
    Else
        'Busca el objeto dentro de la colección
        For i = mcolMenuReporte.Count To 1 Step -1
            Set obj = mcolMenuReporte.item(i)
            If obj Is ix Then       'Si encuentra
                mcolMenuReporte.Remove i       'Elimina el objeto
                GoTo eliminado
            End If
        Next i
        'Si no encuentra el objeto, genera error
        Err.Raise ERR_INVALIDO, "Permiso_RemoveMenuReporte", "No se puede eliminar menu reporte."
    End If
eliminado:
    mProps.objGrupo.SetModificado
    Set obj = Nothing
End Sub

'Duplica un Menu para el mismo Grupo
Public Function CopiaMenuReporte(ix As Long) As Long
    Dim nuevo As PermisoMenuReporte, obj As PermisoMenuReporte
    
    Set obj = mcolMenuReporte.item(ix)
    Set nuevo = obj.Clone
    mcolMenuReporte.Add nuevo
    
    CopiaMenuReporte = mcolMenuReporte.Count       'Devuelve el indice del nuevo elemento
    Set nuevo = Nothing
    Set obj = Nothing
    mProps.objGrupo.SetModificado
End Function

Public Property Get CountMenuReporte() As Long
    CountMenuReporte = mcolMenuReporte.Count
End Property

Public Property Get MenuReporte(ix As Long) As PermisoMenuReporte
    Set MenuReporte = mcolMenuReporte.item(ix)
End Property

Public Sub LimpiaMenuReporte()
    Dim i As Long
    
    For i = mcolMenuReporte.Count To 1 Step -1
        mcolMenuReporte.Remove i
    Next i
End Sub

Friend Sub AddMenuReporteParaClone(obj As PermisoMenuReporte)
    Set obj.Permiso = Me
    mcolMenuReporte.Add obj
End Sub

'AUC 20/07/2010
Public Property Let ConsRepVen(ByVal value As Boolean)
    mProps.ConsRepVen = value
    mProps.objGrupo.SetModificado
End Property

Public Property Get ConsRepVen() As Boolean
    ConsRepVen = mProps.ConsRepVen
End Property

Public Function ListaTransElectronica() As Variant
    Dim i As Long, s As String, v() As Variant, cnt As Integer
    Dim pt As PermisoTrans, sql As String, rs As Recordset
    
        
'            Set pt = mcolTrans.item(i)
            sql = "SELECT codtrans, NombreTrans FROM GNTrans WHERE bandvalida=1 and substring(Opcion2, 122, 1) = 'S' "
            
#If DAOLIB Then
            Set rs = gobjMain.EmpresaActual.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.Open sql, gobjMain.EmpresaActual.Coneccion, adOpenStatic, adLockReadOnly
#End If
            If Not rs.EOF Then
            
                ReDim Preserve v(1, rs.RecordCount)
                For i = 0 To rs.RecordCount - 1
                v(0, cnt) = rs!CodTrans
                v(1, cnt) = rs!NombreTrans
                cnt = cnt + 1
                rs.MoveNext
                Next i
            End If
            rs.Close

        
        ListaTransElectronica = v
        Set pt = Nothing
        Set rs = Nothing
    
End Function

Public Function ListaTransElectronicaLote() As Variant
    Dim i As Long, s As String, v() As Variant, cnt As Integer
    Dim pt As PermisoTrans, sql As String, rs As Recordset
    
        
'            Set pt = mcolTrans.item(i)
            sql = "SELECT codtrans, NombreTrans FROM GNTrans WHERE bandvalida=1 and  substring(Opcion2, 122, 1) = 'S' "
            
#If DAOLIB Then
            Set rs = gobjMain.EmpresaActual.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.Open sql, gobjMain.EmpresaActual.Coneccion, adOpenStatic, adLockReadOnly
#End If
            If Not rs.EOF Then
            
                ReDim Preserve v(1, rs.RecordCount)
                For i = 0 To rs.RecordCount - 1
                v(0, cnt) = rs!CodTrans
                v(1, cnt) = rs!NombreTrans
                cnt = cnt + 1
                rs.MoveNext
                Next i
            End If
            rs.Close

        
        ListaTransElectronicaLote = v
        Set pt = Nothing
        Set rs = Nothing
    
End Function

Public Function ListaTransRecuperaXMLBD() As Variant
    Dim i As Long, s As String, v() As Variant, cnt As Integer
    Dim pt As PermisoTrans, sql As String, rs As Recordset
    
        

            sql = "select transid,claveacceso,archivoxml from infocomprobantes where codigocomprobante =1 order by transid"
            
#If DAOLIB Then
            Set rs = gobjMain.EmpresaActual.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.Open sql, gobjMain.EmpresaActual.Coneccion, adOpenStatic, adLockReadOnly
#End If
            If Not rs.EOF Then
            
                ReDim Preserve v(1, rs.RecordCount)
                For i = 0 To rs.RecordCount - 1
                v(0, cnt) = rs!CodTrans
                v(1, cnt) = rs!NombreTrans
                cnt = cnt + 1
                rs.MoveNext
                Next i
            End If
            rs.Close

        
        ListaTransRecuperaXMLBD = v
        Set pt = Nothing
        Set rs = Nothing
    
End Function


