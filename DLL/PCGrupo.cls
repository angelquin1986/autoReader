VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    NumGrupo As Byte            '1,2,3
    IdGrupo As Long
    CodGrupo As String
    Descripcion As String
    BandValida As Boolean
    Origen As Byte    '1 proveedor, 2 clientes, 3 garante, 4 empleado
    NumeroPagos As Integer
    Intervalo As Integer
    BandDias As Boolean
    idGasto As Long 'AUC enlaza a gngasto
    fechagrabado As Date
    CodGasto As String
    IdTurno As Byte
    Codturno As String

    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean
Private mcolDetallePcGrupo As Collection  'Coleccion de detalles proveedor AUC 22/11/05

Public Sub Grabar()
    Dim sql As String, rs As Recordset
        Dim rsSuc As Recordset, gnSuc As GNSucursal, i As Integer
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
    sql = "SELECT " & tabla("Cod") & " FROM " & tabla("PC") & _
          " WHERE (" & tabla("Cod") & "='" & mProps.CodGrupo & "') AND " & _
                 "(" & tabla("Id") & "<>" & mProps.IdGrupo & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "PCGrupo_Grabar", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM " & tabla("PC") & " WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM " & tabla("PC") & " WHERE " & tabla("Id") & "=" & mProps.IdGrupo
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "PCGrupo_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en Pos
    With rs
        .Fields(tabla("Cod")) = mProps.CodGrupo
        !Descripcion = mProps.Descripcion
        !BandValida = mProps.BandValida
        !Origen = mProps.Origen
        !NumeroPagos = mProps.NumeroPagos
        !Intervalo = mProps.Intervalo
        !BandDias = mProps.BandDias
        !idGasto = mProps.idGasto
        !IdTurno = mProps.IdTurno
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IdGrupo = .Fields(tabla("Id"))
        .Close
    End With
    
    'actualiza Inventario en Bases de Datos Sucursales
    sql = "SELECT * FROM GnSucursal WHERE bandvalida=1"
    Set rsSuc = mProps.objEmpresa.OpenRecordset(sql)
    
    For i = 1 To rsSuc.RecordCount
        Set gnSuc = mProps.objEmpresa.RecuperaGNSucursal(rsSuc.Fields("CodSucursal"))
        If gnSuc.BandCliente Or gnSuc.BandProveedor Then
            GrabarEnSucursales gnSuc.BaseDatos, gnSuc.servidor
        End If
        rsSuc.MoveNext
    Next i
    Set gnSuc = Nothing
    Set rsSuc = Nothing
    GrabarIVDetallePCGrupo
    
        'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_GRU", "Grupo Nuevo: " & _
                                                         mProps.CodGrupo, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_GRU", "Grupo Modificado: " & _
                                                         mProps.CodGrupo, "DT"
    End If
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PCGRUPO1 + NumGrupo - 1
End Sub
Private Function tabla(prefix As String) As String
    tabla = prefix & "Grupo" & NumGrupo
End Function

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM " & tabla("PC") & " WHERE " & tabla("Cod") & "='" & key & "'"
    Else
        sql = "SELECT * FROM " & tabla("PC") & " WHERE " & tabla("Id") & "=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
    'Recupera los datos de Usuario
     With rs
         mProps.IdGrupo = .Fields(tabla("Id"))
         mProps.CodGrupo = .Fields(tabla("Cod"))
         If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
         If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
         If Not IsNull(!Origen) Then mProps.Origen = !Origen
         If Not IsNull(!NumeroPagos) Then mProps.NumeroPagos = !NumeroPagos
         If Not IsNull(!Intervalo) Then mProps.Intervalo = !Intervalo
         If Not IsNull(!BandDias) Then mProps.BandDias = !BandDias
         If Not IsNull(!idGasto) Then mProps.idGasto = !idGasto
        If Not IsNull(!IdTurno) Then mProps.IdTurno = !IdTurno
         
         mProps.fechagrabado = !fechagrabado
         .Close
     End With
             CargaDetallePGGrupo
    End If
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    
    'Verifica si ya existe un item usando el mismo CodGrupo
    sql = "SELECT CodProvCli FROM PCProvCli " & _
          " WHERE " & tabla("Id") & "=" & mProps.IdGrupo
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "PCGrupo_Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close

    sql = "SELECT CodProvCli FROM Empleado " & _
          " WHERE " & tabla("Id") & "=" & mProps.IdGrupo
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "PCGrupo_Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close

    'Elimina el registro
    sql = "DELETE FROM " & tabla("PC") & " WHERE " & tabla("Id") & "=" & mProps.IdGrupo
    mProps.objEmpresa.Execute sql, True

    Set rs = Nothing
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PCGRUPO1 + NumGrupo - 1
End Sub

Friend Function Clone() As PcGrupo
    Dim nuevo As PcGrupo
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New PcGrupo
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdGrupo = 0
    mProps.Origen = 0
    mProps.NumeroPagos = 1
    mProps.Intervalo = 0
    mProps.IdTurno = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property


Friend Property Let NumGrupo(ByVal value As Byte)
    mProps.NumGrupo = value
    mbooModificado = True
End Property

Public Property Get NumGrupo() As Byte
    NumGrupo = mProps.NumGrupo
End Property


Public Property Let CodGrupo(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCGrupo_CodGrupo", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodGrupo = value
    mbooModificado = True
End Property

Public Property Get CodGrupo() As String
    CodGrupo = Trim$(mProps.CodGrupo)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "PCGrupo_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let BandValida(ByVal value As Boolean)
    mProps.BandValida = value
    mbooModificado = True
End Property

Public Property Get BandValida() As Boolean
    BandValida = mProps.BandValida
End Property
Public Property Get fechagrabado() As Date
    fechagrabado = mProps.fechagrabado
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
    
    'Valores iniciales
    mProps.BandValida = True
    mProps.BandDias = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Sub GrabarEnSucursal(ByVal bd As String)
End Sub

Public Sub GrabarEnSucursales(ByVal bd As String, ByVal servidor As String)
   Dim sql As String, rs As Recordset
    
    If mbooEsNuevo Then
        sql = "SELECT " & tabla("Cod") & " FROM " & bd & ".dbo." & tabla("PC") & _
              " WHERE (" & tabla("Cod") & "='" & mProps.CodGrupo & "')  "
        Set rs = mProps.objEmpresa.OpenRecordsetSUC(sql, bd, servidor)

        If Not rs.EOF Then      'Si ya existe el codigo
            Err.Raise ERR_REPITECODIGO, "PCGrupo_Grabar", MSGERR_REPITECODIGO
        End If
        rs.Close
    End If
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM " & bd & ".dbo." & tabla("PC") & " WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEditSuc(sql, bd, servidor)
       rs.AddNew
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM " & bd & ".dbo." & tabla("PC") & " WHERE " & tabla("Id") & "=" & mProps.IdGrupo
        Set rs = mProps.objEmpresa.OpenRecordsetParaEditSuc(sql, bd, servidor)
       If rs.RecordCount = "0" Then
            sql = "SELECT * FROM " & bd & ".dbo." & tabla("PC") & " WHERE 1=0"
            Set rs = mProps.objEmpresa.OpenRecordsetParaEditSuc(sql, bd, servidor)
           rs.AddNew
        End If
#If DAOLIB Then
        rs.Edit
#End If
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
    End If
    'Graba el registro en Pos
    With rs
        .Fields(tabla("Cod")) = mProps.CodGrupo
        !Descripcion = mProps.Descripcion
        !BandValida = mProps.BandValida
        !Origen = mProps.Origen
        !NumeroPagos = mProps.NumeroPagos
        !Intervalo = mProps.Intervalo
        !BandDias = mProps.BandDias
              !IdTurno = mProps.IdTurno
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
'        mProps.IdGrupo = .Fields(tabla("Id"))
        .Close
    End With
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccionEnSucursales bd, servidor, "NUE_GRU", "Grupo Nuevo: " & _
                                                        mProps.CodGrupo, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccionEnSucursales bd, servidor, "MOD_GRU", "Grupo Modificado: " & _
                                                        mProps.CodGrupo, "DT"
    End If
    Set rs = Nothing
End Sub

Public Property Let Modificado(value As Boolean)
    mbooModificado = value
End Property

Public Sub GrabarIVDetallePCGrupo()
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As IVDetallePcGrupo, sql As String, band As Boolean
    Dim NumHijos As Integer, CountNumHijos As Integer
    On Error GoTo CapturaError
    'Elimina  todos  los registros
    sql = "Delete PCGIV " & _
          "From  PCGrupoInventario PCGIV " & _
          "INNER JOIN IVInventario IV  ON iv.Idinventario = PCGIV.Idinventario"
          sql = sql & " Where PCGIV.IdpcGRUPO1 = " & mProps.IdGrupo
    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    band = True
    If Not mcolDetallePcGrupo Is Nothing Then
        sql = "SELECT * FROM PCGrupoInventario WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        NumHijos = mcolDetallePcGrupo.Count
        CountNumHijos = 0
        For i = 1 To mcolDetallePcGrupo.Count
            If CountNumHijos < NumHijos Then
                CountNumHijos = CountNumHijos + 1
                Set obj = mcolDetallePcGrupo.item(i)
                If obj.IdGrupo1 = 0 Then
                    obj.Recuperar obj.CodInventario
                    If verificacionDuplicado(obj.IdInventario) Then
                        RemoveDetallePcGrupo i
                        i = i - 1
                    End If
                End If
                If band Then
                    rs1.AddNew
                    rs1!IdInventario = Me.RecuperaID(obj.CodInventario)
                    rs1!idpcgrupo1 = mProps.IdGrupo
                    rs1!idpcgrupo2 = mProps.IdGrupo
                    rs1!idpcgrupo3 = mProps.IdGrupo
                    rs1!idpcgrupo4 = mProps.IdGrupo
                    rs1!Cantidad = obj.Cantidad
                    rs1!PU = obj.PU
                    rs1!fechagrabado = Now
                    rs1.Update
                End If
            End If
            band = True
        Next i
    End If
    Set rs1 = Nothing
    Set obj = Nothing
    Exit Sub
CapturaError:
    MsgBox Err.Description
    Set rs1 = Nothing
    Set obj = Nothing
    Exit Sub
End Sub

Private Function verificacionDuplicado(ByVal IdInventario As String) As Boolean
    Dim rs As Recordset, sql As String
    verificacionDuplicado = False
    'Recupera los detalles de existencia
    sql = "SELECT IdInventario FROM PCGrupoInventario where idinventario =" & IdInventario
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        verificacionDuplicado = True
    End If
    Set rs = Nothing
End Function

Public Sub RemoveDetallePcGrupo(ByVal ix As Long, Optional obj As IVDetallePcGrupo)
    If mcolDetallePcGrupo Is Nothing Then CargaDetallePGGrupo
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolDetallePcGrupo.Count
            If mcolDetallePcGrupo.item(ix) Is obj Then
                mcolDetallePcGrupo.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolDetallePcGrupo.Remove ix
    End If
    'mcolDetalleFamilia.Remove ix
    mbooModificado = True
End Sub

Private Sub CargaDetallePGGrupo()
    Dim rs As Recordset, sql As String, obj As IVDetallePcGrupo
    Dim i As Long
    Set mcolDetallePcGrupo = New Collection
       sql = "Select IV.IdInventario, Iv.CodInventario, Iv.Descripcion,pcginv.cantidad,pcginv.PU " & _
              "  From VWIvinventarioRECUPERAR Iv " & _
            "INNER JOIN PcGrupoInventario pcgInv "
        sql = sql & " ON pcgInv.IdInventario = IV.Idinventario"
        sql = sql & " Where pcgInv.Idpcgrupo" & NumGrupo & "  = " & mProps.IdGrupo
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddDetallePcGrupo
            Set obj = mcolDetallePcGrupo.item(i)
                  obj.IdInventario = !IdInventario
                  obj.CodInventario = !CodInventario
                  obj.Descripcion = !Descripcion
                  obj.Cantidad = !Cantidad
                  obj.PU = !PU
           .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
        Set obj = Nothing
End Sub

Public Function AddDetallePcGrupo() As Long
    Dim nuevo As IVDetallePcGrupo
    If mcolDetallePcGrupo Is Nothing Then CargaDetallePGGrupo
    Set nuevo = New IVDetallePcGrupo
    Set nuevo.PcGrupo = Me
    nuevo.Cantidad = 1
    mcolDetallePcGrupo.Add nuevo
    AddDetallePcGrupo = mcolDetallePcGrupo.Count    'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function

Public Function RecuperaDetallePcGrupos(ix As Long) As IVDetallePcGrupo
    Set RecuperaDetallePcGrupos = mcolDetallePcGrupo.item(ix)
End Function

Public Property Get NumPCGrupoDetalle() As Long
    If mcolDetallePcGrupo Is Nothing Then
        NumPCGrupoDetalle = 0
    Else
        NumPCGrupoDetalle = mcolDetallePcGrupo.Count
    End If
End Property

Public Function RecuperaID(ByRef cod As String) As Long
    Dim sql As String, rs As Recordset
        sql = "SELECT idinventario FROM ivinventario WHERE Codinventario='" & cod & "'"
    Set rs = Me.Empresa.OpenRecordset(sql)
    With rs
     If rs.RecordCount = 0 Then
        RecuperaID = 0
      Else
        RecuperaID = !IdInventario
        Exit Function
        .Close
      End If
    End With
    Set rs = Nothing
End Function

Public Property Let Origen(ByVal value As Byte)
        mProps.Origen = value
End Property

Public Property Get Origen() As Byte
    Origen = mProps.Origen
End Property

Public Function ListaCuentas() As Recordset
    Dim sql As String, rs As Recordset
    Dim idQueGrupo As Long
    idQueGrupo = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ConAsiento") + 1
    sql = "SELECT pc.idGrupo" & idQueGrupo & ",ele.codElemento,cd.IdCuenta, space(5) As CodCuenta, " & _
          "space(5) As Cuenta  " & _
          "FROM CuentaPcGrupo cd INNER JOIN Elemento ele on ele.idElemento = cd.idElemento " & _
          " INNER JOIN PCGrupo" & idQueGrupo & " pc on pc.idgrupo" & idQueGrupo & " = cd.idpcgrupo " & _
          "WHERE pc.codgrupo" & idQueGrupo & " = '" & Me.CodGrupo & "' ORDER BY ele.orden"
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    Set ListaCuentas = rs
    Set rs = Nothing
End Function

'Eliminar todas las cuentas de un departamento para luego grabar de nuevo
Public Sub EliminaCuentas()
    Dim sql As String, rs As Recordset
        sql = "DELETE FROM CuentaPCGrupo WHERE Idpcgrupo = '" & Me.IdGrupo & "'"
        Me.Empresa.Coneccion.Execute sql
End Sub
'
'Public Function AgregarCuentasDepartamento() As CuentasDepartamento
'    Dim obj As CuentasDepartamento
'    Set obj = New CuentasDepartamento
'    Set obj.Dep = Me
'    Set AgregarCuentasDepartamento = obj
'    Set obj = Nothing
'End Function

Public Property Let NumeroPagos(ByVal value As Integer)
        mProps.NumeroPagos = value
        mbooModificado = True
End Property

Public Property Get NumeroPagos() As Integer
    NumeroPagos = mProps.NumeroPagos
End Property

Public Property Let Intervalo(ByVal value As Integer)
        mProps.Intervalo = value
        mbooModificado = True
End Property

Public Property Get Intervalo() As Integer
    Intervalo = mProps.Intervalo
    
End Property

Public Property Let BandDias(ByVal value As Boolean)
    mProps.BandDias = value
    mbooModificado = True
End Property

Public Property Get BandDias() As Boolean
    BandDias = mProps.BandDias
End Property
Public Property Let IdGrupo(ByVal value As Long)
        mProps.IdGrupo = value
End Property
Public Property Get IdGrupo() As Long
    IdGrupo = mProps.IdGrupo
End Property


Public Function ListaCuentasPre() As Recordset
    Dim sql As String, rs As Recordset
    Dim idQueGrupo As Long
    idQueGrupo = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ConAsiento") + 1
    sql = "SELECT pc.idGrupo" & idQueGrupo & ",ele.codElemento,cd.IdCuenta, space(5) As CodCuenta, " & _
          "space(5) As Cuenta  " & _
          " " & _
          "FROM CuentaPcGrupoPre cd INNER JOIN Elemento ele on ele.idElemento = cd.idElemento " & _
          " INNER JOIN PCGrupo" & idQueGrupo & " pc on pc.idgrupo" & idQueGrupo & " = cd.idpcgrupo " & _
          "WHERE pc.codgrupo" & idQueGrupo & " = '" & Me.CodGrupo & "' ORDER BY ele.orden"
        Set rs = Me.Empresa.OpenRecordset(sql)
    Set ListaCuentasPre = rs
    Set rs = Nothing
End Function

'
Public Function AgregarCuentasPCGrupo() As CuentasPcGrupo
    Dim obj As CuentasPcGrupo
    Set obj = New CuentasPcGrupo
    Set obj.objPCG = Me
    Set AgregarCuentasPCGrupo = obj
    Set obj = Nothing
End Function

Public Sub EliminaCuentasPRE()
    Dim sql As String, rs As Recordset
        sql = "DELETE FROM CuentaPCGrupoPre WHERE Idpcgrupo = '" & Me.IdGrupo & "'"
        Me.Empresa.Coneccion.Execute sql
End Sub

Public Function AgregarCuentasPCGrupoPre() As CuentasPcGrupoPre
    Dim obj As CuentasPcGrupoPre
    Set obj = New CuentasPcGrupoPre
    Set obj.objPCG = Me
    Set AgregarCuentasPCGrupoPre = obj
    Set obj = Nothing
End Function

Public Function AgregarCuentasDepartamento() As CuentasDepartamento
'    Dim obj As CuentasDepartamento
'    Set obj = New CuentasDepartamento
'    Set obj.Dep = Me
'    Set AgregarCuentasDepartamento = obj
'    Set obj = Nothing
End Function

Public Property Let idGasto(ByVal value As Long)
        mProps.idGasto = value
End Property

Public Property Get idGasto() As Long
    idGasto = mProps.idGasto
End Property


Public Property Let CodGasto(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodGasto Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaCosto también
            sql = "SELECT IdGasto FROM GNGasto WHERE CodGasto='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCGrupo_CodGasto", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                idGasto = rs!idGasto
            End If
            rs.Close
        Else
            idGasto = 0
        End If
        mProps.CodGasto = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property

Public Property Get CodGasto() As String
    'CodGasto = Trim$(mProps.CodGasto)
    Dim sql As String, rs As Recordset
    
            sql = "SELECT CodGasto FROM GNGasto WHERE idGasto=" & mProps.idGasto
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                CodGasto = ""
                Exit Property
            Else
                CodGasto = Trim$(rs!CodGasto)
            End If
            rs.Close
    Set rs = Nothing
End Property

 Public Property Let IdTurno(ByVal value As Byte)
        mProps.IdTurno = value
End Property

Public Property Get IdTurno() As Byte
    IdTurno = mProps.IdTurno
End Property

Public Property Let Codturno(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.Codturno Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaCosto también
            sql = "SELECT IdTurno FROM RTurno WHERE CodTurno='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCGrupo_Codturno", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdTurno = rs!IdTurno
            End If
            rs.Close
        Else
            IdTurno = 0
        End If
        mProps.Codturno = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property

Public Property Get Codturno() As String
    'CodGasto = Trim$(mProps.CodGasto)
    Dim sql As String, rs As Recordset
            sql = "SELECT CodTurno FROM RTurno WHERE idturno=" & mProps.IdTurno
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Codturno = ""
                Exit Property
            Else
                Codturno = Trim$(rs!Codturno)
            End If
            rs.Close
    Set rs = Nothing
End Property


