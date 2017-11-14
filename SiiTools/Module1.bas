Attribute VB_Name = "Module1"
Option Explicit

Private Type T_ConfigIVFisico '***Angel. 19/mar/04
    CodTrans_CF As String
    CodTrans_AJ As String
    CodTrans_BJ As String
    BandLineaAuto As Boolean
    BandTotalizarItem As Boolean    'jeaa 13/10/04
End Type

Private Type T_ConfigIVAjusteAutomatico '***Esteban 20/09/2006
    CodTrans_AA As String
    CodTrans_AAJ As String
    CodTrans_ABJ As String
    BandLineaAutoA As Boolean
    BandTotalizarItemA As Boolean    'jeaa 13/10/04
End Type

Type Comisiones '08/02/2006 tabla para comisions
    desde As Currency
    hasta As Currency
    Comision As Currency
    ComisionC As Currency
    ComisionSC As Currency
End Type

Type Config
    'Archivo en donde se encuentre las fórmulas para el calculo de comisiones
    Archivo As String
    ArchivoB As String
    ArchivoC As String
    ArchivoD As String
    PorcenVendedorA As Currency
    PorcenVendedorB As Currency
    PorcenVendedorC As Currency
    PorcenVendedorD As Currency
    PorcenCobradorA As Currency
    PorcenCobradorB As Currency
    PorcenCobradorC As Currency
    PorcenCobradorD As Currency
End Type


Type ConfigJefe
    'Archivo en donde se encuentre las fórmulas para el calculo de comisiones
    ArchivoJefeA As String
    ArchivoJefeB As String
    ArchivoJefeC As String
    ArchivoJefeD As String
End Type

Type Montos
    desde As Currency
    hasta As Currency
    grupo As String
End Type

Type MontosCobros
    desde As Currency
    hasta As Currency
    diasMorosidad As Currency
    grupo As String
End Type



Public Const NOCONTADO = "Item no contado fisicamente" '***Angel. 10/sep/2004
Public Const COLOR_NOCONTADO = vbBlue                  '***Angel. 10/sep/2004

'Anchos de columnas
Public Const COLANCHO_CUR = 1400
Public Const COLANCHO_FECHA = 1200
Public Const COLANCHO_CANT = 1000
Public Const ERR_IVFILTROBODEGA = ERRNUM + 23
Public Const MSGERR_IVFILTROBODEGA = "No grabada por filtro de bodega"

Type T_CONFIG        'configuraciones para importar transaccions venta de locutorios
    CodTrans  As String
    CodCli As String
    MONEDA As String
    Responsable As String
    FormaCobroPago As String
    AbrirArchivoenFormaDiferencial As Boolean
End Type

'***Angel. 22/Mar/2004
'Public Const APPNAME_HIDE = "CLSID\{1B7C788B-E925-438F-88C4-FDCF166BF53D_10}\PROGID"

'Variables globales para todo el programa
Public gobjMain As SiiMain
Public gConfig As T_CONFIG
Public Const ARCHIVO_MODELO = "_Trans.sys"
Public gConfigIVFisico As T_ConfigIVFisico '***Angel. 19/mar/04
Public gConfigIVAjusteAutomatico As T_ConfigIVAjusteAutomatico '*** jeaa 20/09/2006


Public gComisiones(1 To 10) As Comisiones
Public gComisionesB(1 To 10) As Comisiones
Public gComisionesC(1 To 10) As Comisiones
Public gComisionesD(1 To 10) As Comisiones
Public gComisionesJefe(1 To 10) As Comisiones
Public gComisionesJefeA(1 To 10) As Comisiones
Public gComisionesJefeB(1 To 10) As Comisiones
Public gComisionesJefeC(1 To 10) As Comisiones
Public gComisionesJefeD(1 To 10) As Comisiones
Public gConfigura As Config
Public gConfiguraJefe As ConfigJefe
Public gobjRol As Object
Public gMonto(1 To 10) As Montos
Public gMontoCobro(1 To 10) As MontosCobros

Public Sub Main()
    Dim code As String, pos As Integer
    Dim BandModulo As Boolean
    On Error GoTo ErrTrap
    
'    frmSplash.Inicio
    
    Set gobjMain = New SiiMain
    gobjMain.Inicializar
    
    Load frmMain
 '   Unload frmSplash

    If Not frmLoginSplash.Inicio Then End
    Unload frmLoginSplash
    frmMain.Show
    'Obtiene codigo de la ultima empresa
    code = gobjMain.EmpresaAnterior
    'Si no puede recuperar, selecciona
    If Len(code) = 0 Then
        frmMain.mnuAbrirEmpresa_Click
    'Si recupera la empresa anterior, abre la misma
    Else
        'Si no la puede abrir, hace seleccionar
        If Not AbrirEmpresa(code, False) Then
            frmMain.mnuAbrirEmpresa_Click
        End If
    End If
    
    frmMain.mnuHerramienta.Enabled = Not (gobjMain.EmpresaActual Is Nothing)
    frmMain.mnuTransferir.Enabled = Not (gobjMain.EmpresaActual Is Nothing)
    
    
    'Sólo supervisores tiene derecho a hacer Exportacion/Importacion        '*** MAKOTO 12/ene/01
'    frmMain.mnuTransferir.Enabled = gobjMain.UsuarioActual.BandSupervisor
'    frmMain.mnuImportar.Enabled = gobjMain.UsuarioActual.BandSupervisor
    frmMain.mnuCierre.Enabled = gobjMain.UsuarioActual.BandSupervisor
    frmMain.mnuCreaTransAnulada.Visible = gobjMain.UsuarioActual.BandSupervisor
    frmMain.mnuDINARDAP.Visible = gobjMain.EmpresaActual.GNOpcion.BandDINARDAP
            
    '***Angel. 23/Marzo/2004
    'frmMain.mnuIVFisico.Visible = frmMain.RecuperaRegistroIVFisico
    
'''    'jeaa 01/04/2007
'''    If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("FormulariosSRI") = "1" Then
'''        frmMain.mnuDeclaracionesI.Visible = True
'''    Else
'''        frmMain.mnuDeclaracionesI.Visible = False
'''    End If
    'jeaa 08/06/2007
    pos = InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "HORMI")
    If pos > 0 Then
        frmMain.mnuCierreTransXEntregarF.Visible = True
    End If
   
    BandModulo = gobjMain.PermisoModuloEspecial(gobjMain.UsuarioActual.codUsuario, ModuloPuntoEqui)
    If BandModulo Then
        If pos > 0 Then
            frmMain.mnuPe.Visible = True
            frmMain.mnupeMP3.Visible = False
        Else
            frmMain.mnuPe.Visible = False
            frmMain.mnupeMP3.Visible = True
        End If
    Else
        frmMain.mnuPe.Visible = False
        frmMain.mnupeMP3.Visible = False
    End If
    'AUC FACTURACION X LOTE
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("FacturarxLoteCliente")) > 0 Then
        If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("FacturarxLoteCliente") = "1" Then
            frmMain.mnuLoteCli.Visible = True
        Else
            frmMain.mnuLoteCli.Visible = False
        End If
    Else
        frmMain.mnuLoteCli.Visible = False
    End If
    
'    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Anexos2015")) > 0 Then
'        If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Anexos2015") = 1 Then
'            frmMain.mnuAnexoTra2015.Visible = True
'        Else
'            frmMain.mnuAnexoTra2015.Visible = False
'        End If
'    Else
'            frmMain.mnuAnexoTra2015.Visible = False
'    End If
    
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Anexos2016")) > 0 Then
        If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Anexos2016") = 1 Then
            frmMain.mnuAnexoTra2016.Visible = True
        Else
            frmMain.mnuAnexoTra2016.Visible = False
        End If
    Else
            frmMain.mnuAnexoTra2016.Visible = False
    End If
    
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Anexos2016C")) > 0 Then
        If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Anexos2016C") = 1 Then
            frmMain.mnuAnexoTra2016Consol.Visible = True
        Else
            frmMain.mnuAnexoTra2016Consol.Visible = False
        End If
    Else
            frmMain.mnuAnexoTra2016Consol.Visible = False
    End If
    If gobjMain.UsuarioActual.BandSupervisor Then
        frmMain.mnuRegColas.Enabled = True
    Else
        frmMain.mnuRegColas.Enabled = False
    End If
    Exit Sub
ErrTrap:
    If Err.Number = ERR_NOREGINFO Then
        MsgBox "El programa se inicia por primera vez." & vbCr & _
               "Llene las configuraciones iniciales del sistema."
    Else
        DispErr
    End If
    End
    Exit Sub
End Sub

'Recibe codigo de empresa y la abre
Public Function AbrirEmpresa(ByVal cod As String, ByVal mensaje As Boolean) As Boolean
    Dim emp As Empresa
    On Error GoTo ErrTrap
    AbrirEmpresa = False
    MensajeStatus "Está abriendo la empresa ...", vbHourglass
    Set emp = gobjMain.RecuperaEmpresa(cod)
    If Not (emp Is Nothing) Then
        If Not (gobjMain.EmpresaActual Is Nothing) Then
            'gobjMain.EmpresaActual.Cerrar
            gobjMain.EmpresaActual.CerrarModulo ModuloTools
        End If
        'Abre la base de datos de la empresa
        'emp.Abrir
        emp.AbrirModulo ModuloTools
        AbrirEmpresa = True
    ElseIf mensaje Then
        MensajeStatus "", 0
        MsgBox "No se puede abrir la empresa '" & cod & "'."
    End If
    Set emp = Nothing
    frmMain.CambiaCaption       'Actualiza la Caption
    MensajeStatus "", 0
    Exit Function
ErrTrap:
    MensajeStatus "", 0
    If mensaje Then DispErr
    Exit Function
End Function

Public Sub MensajeStatus(Optional msg As String, Optional puntero As Integer)
    With frmMain
        Screen.MousePointer = puntero        'Cambia la figura de mouse
        .stb1.Panels("msg").Text = msg
    End With
End Sub


Public Sub GNPoneNumFila(grd As Object, PonerEnSubtotal As Boolean)
'    Dim i As Long
'
'    With grd
'        For i = .FixedRows To .Rows - 1
'            If (Not .IsSubtotal(i)) Or PonerEnSubtotal Then .TextMatrix(i, 0) = i
'        Next i
'    End With
    
        Dim i As Long, contador As Long
    
    With grd
        contador = .FixedRows
        For i = .FixedRows To .Rows - 1
            If (Not .IsSubtotal(i)) Or PonerEnSubtotal Then
                .TextMatrix(i, 0) = contador
                contador = contador + 1
            End If
        Next i
    End With

    
End Sub

Public Sub SeleccionaComboItem(cbo As ComboBox, cod As String)
    Dim i As Integer
    
    cbo.ListIndex = -1
    For i = 0 To cbo.ListCount - 1
        If cbo.List(i) = cod Then
            cbo.ListIndex = i
        End If
    Next i
End Sub

'Busca formulario cargado en la memoria para que no se genere dos instancias de la misma ventana
Public Function BuscaForm(ByVal Name As String, ByVal tag As String) As Form
    Dim frm As Form
    
    For Each frm In Forms
        If (UCase$(frm.Name) = UCase$(Name)) And (UCase$(frm.tag) = UCase$(tag)) Then
            Set BuscaForm = frm
            Exit For
        End If
    Next frm
    Set frm = Nothing
End Function



Public Sub CargarCatalogos( _
                ByVal grd As VSFlexGrid)
    Dim i As Integer

    With grd
        .Cols = 3
        .Rows = .FixedRows
        .AddItem .Rows & vbTab & "Responsables" & vbTab & "GNResp"
        .AddItem .Rows & vbTab & "Plan de cuenta" & vbTab & "CTCuenta"

        .AddItem .Rows & vbTab & "Bancos" & vbTab & "TSBanco"
        .AddItem .Rows & vbTab & "Retenciones" & vbTab & "TSRetencion"  '*** MAKOTO 12/feb/01 Agregado
        .AddItem .Rows & vbTab & "Bodegas" & vbTab & "IVBodega"
        
        For i = 1 To IVGRUPO_MAX
            .AddItem .Rows & vbTab & _
                gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(i) & " de inventario" & _
                vbTab & "IVG" & i
        Next i
        
        .AddItem .Rows & vbTab & "Vendedores" & vbTab & "FCVendedor"
    
        For i = 1 To PCGRUPO_MAX
            .AddItem .Rows & vbTab & _
                gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupo(i) & " de proveedor/cliente" & _
                vbTab & "PCG" & i
        Next i
        
        .AddItem .Rows & vbTab & "DiasCred" & vbTab & "DiasCred"
        .AddItem .Rows & vbTab & "Parroquias" & vbTab & "PCParroquia"
        .AddItem .Rows & vbTab & "Proveedores" & vbTab & "PCProvCli(P)"
        .AddItem .Rows & vbTab & "Clientes" & vbTab & "PCProvCli(C)"
        .AddItem .Rows & vbTab & "Garantes" & vbTab & "PCProvCli(G)"
        '*** Cambio Oliver 12/05/2003
        .AddItem .Rows & vbTab & "Centro de costo" & vbTab & "GNCentroCosto"
        '**** jeaa 17/07/2006
        .AddItem .Rows & vbTab & "IVUnidad" & vbTab & "IVU"
        .AddItem .Rows & vbTab & "Inventarios" & vbTab & "IVInv"     'Debe estar después de Proveedor
        '**** jeaa 04/01/2005
        .AddItem .Rows & vbTab & "Descuentos PCGrupo x IVGrupo" & vbTab & "DescIVGPCG"     'Debe estar después de PCGRUPOs e IVGRUPOS
        .AddItem .Rows & vbTab & "Forma de Cobro/Pago" & vbTab & "TSFormaC_P"
        '**** jeaa 17/05/2005
        .AddItem .Rows & vbTab & "Motivos Devolucion" & vbTab & "Motivo"     'Debe estar después de PCGRUPOs e IVGRUPOS
        '**** jeaa 26/12/2005
        .AddItem .Rows & vbTab & "Tipo Compras" & vbTab & "TCompra"
        'AUC 13/11/07
        .AddItem .Rows & vbTab & "Historial de Clientes" & vbTab & "PCHistorial"
        '**** jeaa 18/02/2008
        .AddItem .Rows & vbTab & "Existencia" & vbTab & "Exist"
        '**** jeaa 12/09/2008
        .AddItem .Rows & vbTab & "Descuentos NumPagos x IVGrupo" & vbTab & "DescNumPagIVG"     'Debe estar después de PCGRUPOs e IVGRUPOS
        '**** jeaa 22/07/2009
        .AddItem .Rows & vbTab & "IVBancos" & vbTab & "IVBanco"
        '**** jeaa 22/07/2009
        .AddItem .Rows & vbTab & "IVTarjeta" & vbTab & "IVTarjeta"
        .AddItem .Rows & vbTab & "Plazo IVGrupo x PCGrupo" & vbTab & "PLAIVGPCG"     'Debe estar después de PCGRUPOs e IVGRUPOS
        
        
        .FormatString = "^#|<Catálogo|<Tabla"
        GNPoneNumFila grd, False
        AsignarTituloAColKey grd            'Para usar ColIndex
        AjustarAutoSize grd, -1, -1, 3000  'Ajusta automáticamente ancho de cols.
        
        'Oculta columnas innecesarias
        .ColHidden(.ColIndex("Tabla")) = True
    End With
End Sub

#If DAOLIB Then
#Else
Public Function RecuperarCampo( _
                ByVal tabla As String, _
                ByVal campo As String, _
                ByVal cond As String) As Variant
    Dim sql As String, rs As Recordset
    
    sql = "SELECT " & campo & " FROM " & tabla
    If Len(cond) > 0 Then
        sql = sql & " WHERE " & cond
    End If
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not rs.EOF Then
        RecuperarCampo = rs.Fields(campo)
    Else
        If rs.Fields(campo).Type = adVarChar _
            Or rs.Fields(campo).Type = adChar Then
            RecuperarCampo = ""
        Else
            RecuperarCampo = 0
        End If
    End If
    rs.Close
    Set rs = Nothing
End Function
#End If



Public Function FlexCodigosSeleccionados( _
                    ByVal grd As VSFlexGrid, _
                    ByVal ColCodigo As Long, _
                    ByVal comilla As Boolean) As String
    Dim i As Long, s As String
    
    With grd
        For i = .FixedRows To .Rows - 1
            If .IsSelected(i) Then
                If Len(s) > 0 Then s = s & ","
                If comilla Then s = s & "'"
                s = s & .TextMatrix(i, ColCodigo)
                If comilla Then s = s & "'"
            End If
        Next i
    End With
    FlexCodigosSeleccionados = s
End Function




Public Function VerificarSeleccionado( _
                    ByVal grd1 As VSFlexGrid, _
                    ByVal grd2 As VSFlexGrid, _
                    ByVal Desc As String) As Boolean
    If grd1.SelectedRows = 0 And grd2.SelectedRows = 0 Then
        If MsgBox("No está seleccionada ningúna fila " & _
                "en Transacciones ni Catálogos." & vbCr & vbCr & _
                "Desea " & Desc & " TODOS los datos mostrados?", vbQuestion + vbYesNo) <> vbYes Then
            MsgBox "Seleccione los datos que que desea " & Desc & " e inténte de nuevo.", vbInformation
            Exit Function
        Else
            If grd1.Rows > grd1.FixedRows Then
                grd1.Select grd1.FixedRows, grd1.FixedCols, grd1.Rows - 1, grd1.Cols - 1
            End If
            If grd2.Rows > grd2.FixedRows Then
                grd2.Select grd2.FixedRows, grd2.FixedCols, grd2.Rows - 1, grd2.Cols - 1
            End If
        End If
    End If
    
    VerificarSeleccionado = True
End Function


Public Sub LimpiarSeleccion(ByVal grd As VSFlexGrid)
    Dim i As Long
    
    With grd
        For i = .FixedRows To .Rows - 1
            .IsSelected(i) = False
        Next i
    End With
End Sub


'Public Sub PrepararGNComprobante( _
'                ByVal gc As GNComprobante, _
'                ByRef Estado As Byte)
'    Dim sql As String, rs As Recordset, id As Long
'
'    'Abre el orígen para recuperar registro
'    sql = "SELECT * FROM GNComprobante " & _
'          "WHERE CodTrans = '" & gc.CodTrans & "' AND NumTrans = " & gc.NumTrans
'    Set rs = New Recordset
'    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
'
'    With gc
''        .CodAsiento = rs.Fields("CodAsiento")
'        .FechaTrans = rs.Fields("FechaTrans")
'        .HoraTrans = rs.Fields("HoraTrans")
'        If Not IsNull(rs.Fields("Descripcion")) Then .Descripcion = rs.Fields("Descripcion")
'        .CodUsuario = rs.Fields("CodUsuario")      'Solo lectura
'        .CodResponsable = rs.Fields("CodResponsable")
'        If Not IsNull(rs.Fields("NumDocRef")) Then .NumDocRef = rs.Fields("NumDocRef")
'        If Not IsNull(rs.Fields("Nombre")) Then .Nombre = rs.Fields("Nombre")
'        Estado = rs.Fields("Estado")        'Devuelve el estado para forzar a grabar con este valor
''        rs.Fields("PosID") = .PosID                'Solo lectura
'        .NumTransCierrePOS = rs.Fields("NumTransCierrePOS")
'        If Not IsNull(rs.Fields("CodCentro")) Then .CodCentro = rs.Fields("CodCentro")
'
'        'CodTransFuente + NumTransFuente --> IdTransFuente
'        id = RecuperarCampo("GNComprobante", "TransID", _
'                    "CodTrans='" & rs.Fields("CodTransFuente") & "' AND " & _
'                    "NumTrans=" & rs.Fields("NumTransFuente"))
'        .IdTransFuente = id
'
'        .CodMoneda = rs.Fields("CodMoneda")
'        .Cotizacion(2) = rs.Fields("Cotizacion2")
'        .Cotizacion(3) = rs.Fields("Cotizacion3")
'        .Cotizacion(4) = rs.Fields("Cotizacion4")
'
'        'Codxxxx --> IdProveedorRef,IdClienteRef, IdVendedor (Hace dentro el objeto)
'        If Not IsNull(rs.Fields("CodProveedorRef")) Then .CodProveedorRef = rs.Fields("CodProveedorRef")
'        If Not IsNull(rs.Fields("CodClienteRef")) Then .CodClienteRef = rs.Fields("CodClienteRef")
'        If Not IsNull(rs.Fields("CodVendedor")) Then .CodVendedor = rs.Fields("CodVendedor")
'
'    End With
'    rs.Close
'
'    Set gc = Nothing
'    Set rs = Nothing
'End Sub
'
'Public Sub PrepararIVKardex(ByVal gc As GNComprobante)
'    Dim sql As String, rs As Recordset, ivk As IVKardex, i As Long
'
'    'Primero limpia
'    gc.BorrarIVKardex
'
'    'Abre el destino para agregar registro
'    sql = "SELECT * FROM IVKardex " & _
'          "WHERE CodTrans = '" & gc.CodTrans & "' AND NumTrans = " & gc.NumTrans & _
'          " ORDER BY Orden"
'    Set rs = New Recordset
'    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
'
'    Do Until rs.EOF
'        DoEvents
'
'        i = gc.AddIVKardex
'        Set ivk = gc.IVKardex(i)
'        With ivk
'            .CodInventario = rs.Fields("CodInventario")
'            .CodBodega = rs.Fields("CodBodega")
'            .Cantidad = rs.Fields("Cantidad")
'            .CostoTotal = rs.Fields("CostoTotal")
'            .CostoRealTotal = rs.Fields("CostoRealTotal")
'            .PrecioTotal = rs.Fields("PrecioTotal")
'            .PrecioRealTotal = rs.Fields("PrecioRealTotal")
'            If Not IsNull(rs.Fields("Descuento")) Then .Descuento = rs.Fields("Descuento")
'            If Not IsNull(rs.Fields("IVA")) Then .IVA = rs.Fields("IVA")
'            .Orden = rs.Fields("Orden")
'            If Not IsNull(rs.Fields("Nota")) Then .Nota = rs.Fields("Nota")
'        End With
'
'        rs.MoveNext
'    Loop
'
'    rs.Close
'    Set gc = Nothing
'    Set ivk = Nothing
'    Set rs = Nothing
'End Sub
'
'Public Sub PrepararIVKardexRecargo(ByVal gc As GNComprobante)
'    Dim sql As String, rs As Recordset, ivkr As IVKardexRecargo, i As Long
'
'    'Primero limpia
'    gc.BorrarIVKardexRecargo
'
'    'Abre el destino para agregar registro
'    sql = "SELECT * FROM IVKardexRecargo " & _
'          "WHERE CodTrans = '" & gc.CodTrans & "' AND NumTrans = " & gc.NumTrans & _
'          " ORDER BY Orden"
'    Set rs = New Recordset
'    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
'
'    Do Until rs.EOF
'        DoEvents
'
'        i = gc.AddIVKardexRecargo
'        Set ivkr = gc.IVKardexRecargo(i)
'        With ivkr
'            .CodRecargo = rs.Fields("CodRecargo")
'            .Porcentaje = rs.Fields("Porcentaje")
'            .Valor = rs.Fields("Valor")
'            .BandModificable = rs.Fields("BandModificable")
'            .BandOrigen = rs.Fields("BandOrigen")
'            .BandProrrateado = rs.Fields("BandProrrateado")
'            .AfectaIvaItem = rs.Fields("AfectaIvaItem")
'            .Orden = rs.Fields("Orden")
'        End With
'
'        rs.MoveNext
'    Loop
'
'    rs.Close
'    Set gc = Nothing
'    Set ivkr = Nothing
'    Set rs = Nothing
'End Sub
'
'Private Sub PrepararPCKardex(ByVal gc As GNComprobante)
'    Dim sql As String, rs As Recordset, pck As PCKardex, i As Long
'    Dim idAsignado As Long
'    Dim v() As String
'
'    'Primero limpia
'    gc.BorrarPCKardex
'
'    'Abre el destino para agregar registro
'    sql = "SELECT * FROM PCKardex " & _
'          "WHERE CodTrans = '" & gc.CodTrans & "' AND NumTrans = " & gc.NumTrans & _
'          " ORDER BY Orden"
'    Set rs = New Recordset
'    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
'
'    Do Until rs.EOF
'        DoEvents
'
'        i = gc.AddPCKardex
'        Set pck = gc.PCKardex(i)
'        With pck
'            'Desactiva la verificación de saldo de doc.asignado
'            'Para que no genere error cuando asigna valor de Debe/Haber
'            .BandNoVerificarSaldo = True            '*** MAKOTO 22/mar/01 Agregado
'
'            .codProvCli = rs.Fields("CodProvCli")
'
'            If Len(rs.Fields("GuidAsignado")) > 0 Then      '*** MAKOTO 16/mar/01
'                .SetIdAsignadoPorGuid rs.Fields("GuidAsignado")
'            End If
'
'            .CodForma = rs.Fields("CodForma")
'            If Not IsNull(rs.Fields("NumLetra")) Then .NumLetra = rs.Fields("NumLetra")
'            .Debe = rs.Fields("Debe")
'            .Haber = rs.Fields("Haber")
'            .FechaEmision = rs.Fields("FechaEmision")
'            .FechaVenci = rs.Fields("FechaVenci")
'            If Not IsNull(rs.Fields("Observacion")) Then .Observacion = rs.Fields("Observacion")
'            .Orden = rs.Fields("Orden")
'
'            '*** MAKOTO 16/mar/01 Agregado
'            .Guid = rs.Fields("guid")
'            .SetIdFromGuid
'        End With
'
'        rs.MoveNext
'    Loop
'
'    rs.Close
'    Set gc = Nothing
'    Set pck = Nothing
'    Set rs = Nothing
'End Sub
'
'Public Sub PrepararTSKardex(ByVal gc As GNComprobante)
'    Dim sql As String, rs As Recordset, tsk As TSKardex, i As Long
'
'    'Primero limpia
'    gc.BorrarTSKardex
'
'    'Abre el destino para agregar registro
'    sql = "SELECT * FROM TSKardex " & _
'          "WHERE CodTrans = '" & gc.CodTrans & "' AND NumTrans = " & gc.NumTrans & _
'          " ORDER BY Orden"
'    Set rs = New Recordset
'    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
'
'    Do Until rs.EOF
'        DoEvents
'
'        i = gc.AddTSKardex
'        Set tsk = gc.TSKardex(i)
'        With tsk
'            .CodBanco = rs.Fields("CodBanco")
'            .Debe = rs.Fields("Debe")
'            .Haber = rs.Fields("Haber")
'            If Not IsNull(rs.Fields("Nombre")) Then .Nombre = rs.Fields("Nombre")
'            .CodTipoDoc = rs.Fields("CodTipoDoc")
'            If Not IsNull(rs.Fields("NumDoc")) Then .numdoc = rs.Fields("NumDoc")
'            .FechaEmision = rs.Fields("FechaEmision")
'            .FechaVenci = rs.Fields("FechaVenci")
'            If Not IsNull(rs.Fields("Observacion")) Then .Observacion = rs.Fields("Observacion")
'            If Not IsNull(rs.Fields("BandConciliado")) Then .BandConciliado = rs.Fields("BandConciliado")
'            .Orden = rs.Fields("Orden")
'        End With
'
'        rs.MoveNext
'    Loop
'
'    rs.Close
'    Set gc = Nothing
'    Set tsk = Nothing
'    Set rs = Nothing
'End Sub
'
''*** MAKOTO 12/feb/01 Agregado
'Public Sub PrepararTSKardexRet(ByVal gc As GNComprobante)
'    Dim sql As String, rs As Recordset, tskr As TSKardexRet, i As Long
'
'    'Primero limpia
'    gc.BorrarTSKardexRet
'
'    'Abre el destino para agregar registro
'    sql = "SELECT * FROM TSKardexRet " & _
'          "WHERE CodTrans = '" & gc.CodTrans & "' AND NumTrans = " & gc.NumTrans & _
'          " ORDER BY Orden"
'    Set rs = New Recordset
'    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
'
'    Do Until rs.EOF
'        DoEvents
'
'        i = gc.AddTSKardexRet
'        Set tskr = gc.TSKardexRet(i)
'        With tskr
'            .CodRetencion = rs.Fields("CodRetencion")
'            .Valor = rs.Fields("Debe") + rs.Fields("Haber")
'            .Base = rs.Fields("Base")
'            If Not IsNull(rs.Fields("NumDoc")) Then .numdoc = rs.Fields("NumDoc")
'            If Not IsNull(rs.Fields("Observacion")) Then .Observacion = rs.Fields("Observacion")
'            .Orden = rs.Fields("Orden")
'        End With
'
'        rs.MoveNext
'    Loop
'
'    rs.Close
'    Set gc = Nothing
'    Set tskr = Nothing
'    Set rs = Nothing
'End Sub
'
'Public Sub PrepararCTLibroDetalle(ByVal gc As GNComprobante)
'    Dim sql As String, rs As Recordset, ctd As CTLibroDetalle, i As Long
'
'    'Primero limpia
'    gc.BorrarCTLibroDetalle
'
'    'Abre el destino para agregar registro
'    sql = "SELECT * FROM CTLibroDetalle " & _
'          "WHERE CodTrans = '" & gc.CodTrans & "' AND NumTrans = " & gc.NumTrans & _
'          " ORDER BY Orden"
'    Set rs = New Recordset
'    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
'
'    Do Until rs.EOF
'        DoEvents
'
'        i = gc.AddCTLibroDetalle
'        Set ctd = gc.CTLibroDetalle(i)
'        With ctd
'            .codcuenta = rs.Fields("CodCuenta")
'            If Not IsNull(rs.Fields("Descripcion")) Then .Descripcion = rs.Fields("Descripcion")
'            .Debe = rs.Fields("Debe")
'            .Haber = rs.Fields("Haber")
'            .BandIntegridad = rs.Fields("BandIntegridad")
'            .Orden = rs.Fields("Orden")
'        End With
'
'        rs.MoveNext
'    Loop
'
'    rs.Close
'    Set gc = Nothing
'    Set ctd = Nothing
'    Set rs = Nothing
'End Sub
'
'*** Oliver funcion agragada para poder imprimir el asiento

Public Function ImprimeAsiento(ByVal gc As GNComprobante, ByRef objImp As Object) As Boolean
    Dim crear As Boolean
    On Error GoTo ErrTrap
    ImprimeAsiento = False
    
    'Si no tiene TransID quere decir que no está grabada
'    If (gc.TransID = 0) Or gc.Modificado Then
    If (gc.TransID = 0) Then            '*** MAKOTO 07/jul/2000
        MsgBox MSGERR_NOGRABADO, vbInformation
        Exit Function
    End If
    
    'Solo por primera vez o cuando cambia la librería de impresión
    '  crea una instancia del objeto para la impresión
    crear = (objImp Is Nothing)
    If Not crear Then crear = (objImp.NombreDLL <> gc.GNTrans.ArchivoReporte)
    If crear Then
        Set objImp = Nothing
        Set objImp = CreateObject(gc.GNTrans.ArchivoReporte & ".PrintTrans")
    End If
    
    MensajeStatus MSG_PREPARA, vbHourglass
    objImp.PrintAsiento gobjMain.EmpresaActual, False, 1, 0, "", 0, gc
    MensajeStatus
    ImprimeAsiento = True
    Exit Function
ErrTrap:
    MensajeStatus
    Select Case Err.Number
    Case ERR_NOIMPRIME, ERR_NOIMPRIME2, ERR_NOIMPRIME3, ERR_NOHAYCODIGO
        DispErr
    Case Else
        MsgBox MSGERR_NOIMPRIME2, vbInformation
    End Select
    ImprimeAsiento = False
    Exit Function
End Function

Public Function VerificaAsiento(gnc As GNComprobante) As Boolean
    Dim i As Long, j As Long, obj As CTLibroDetalle
    Dim verCuadrado As Boolean, verIntegridad As Boolean
    Dim msg As String
    On Error GoTo ErrTrap

    verCuadrado = True
    verIntegridad = True

OtraVez:
    gnc.VerificaAsiento verCuadrado, verIntegridad
    VerificaAsiento = True
    Exit Function
ErrTrap:
    Select Case Err.Number
    Case ERR_DESCUADRADO
        msg = "El asiento no está cuadrado, por lo que " & _
              "no podrá ser aprobado hasta que esté cuadrado." & vbCr & vbCr & _
              "  Debe: " & Format(gnc.DebeTotal, "#,0.0000") & _
              "  Haber: " & Format(gnc.HaberTotal, "#,0.0000") & _
              "  Diferencia: " & Format(gnc.DebeTotal - gnc.HaberTotal, "#,0.0000") & _
              " " & gnc.CodMoneda & vbCr & vbCr & _
              "Desea continuar para cuadrarlo después?"
              
        If MsgBox(msg, vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
            verCuadrado = False
            Resume OtraVez
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
    Case ERR_INTEGRIDAD
        msg = MSGERR_INTEGRIDAD2 & vbCr & vbCr
        For i = 1 To gnc.CountCTLibroDetalle
            Set obj = gnc.CTLibroDetalle(i)
            If (obj.BandIntegridad <> INTEG_INTEGRADO) And _
               (obj.BandIntegridad <> INTEG_AUTO) Then
                msg = msg & Format(i, "(#)  ") & obj.codcuenta & " " & obj.auxNombreCuenta & vbCr
            End If
        Next i
        Set obj = Nothing
        msg = msg & vbCr & "Desea continuar de todas maneras para después revisar y modificar?"
        If MsgBox(msg, vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
            verIntegridad = False
            Resume OtraVez
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
    Case Else
        DispErr
    End Select
    Exit Function
End Function

Public Function NullSiZero(v As Variant) As Variant
    If v = 0 Then
        NullSiZero = Null
    Else
        NullSiZero = v
    End If
End Function

Public Sub RecuperarConfigIVFisico() '***Angel. 19/mar/04
    With gConfigIVFisico
        .CodTrans_CF = GetSetting(APPNAME, App.Title, "CodTrans_CF", "OF")
        .CodTrans_AJ = GetSetting(APPNAME, App.Title, "CodTrans_AJ", "JI")
        .CodTrans_BJ = GetSetting(APPNAME, App.Title, "CodTrans_BJ", "BD")
        .BandLineaAuto = GetSetting(APPNAME, App.Title, "BandLineaAuto", False)
        'jeaa 13/10/04
        .BandTotalizarItem = GetSetting(APPNAME, App.Title, "BandTotalizarItem", False)
    End With
End Sub

Public Sub GrabarConfigIVFisico() '***Angel. 19/mar/04
    With gConfigIVFisico
        SaveSetting APPNAME, App.Title, "CodTrans_CF", .CodTrans_CF
        SaveSetting APPNAME, App.Title, "CodTrans_AJ", .CodTrans_AJ
        SaveSetting APPNAME, App.Title, "CodTrans_BJ", .CodTrans_BJ
        SaveSetting APPNAME, App.Title, "BandLineaAuto", .BandLineaAuto
        'jeaa 13/10/04
        SaveSetting APPNAME, App.Title, "BandTotalizarItem", .BandTotalizarItem
    End With
End Sub

Public Sub RecuperaConfig()
    On Error GoTo ErrTrap
    With gConfig
        .CodCli = GetSetting(APPNAME, App.Title, "CodCli", "")
        .CodTrans = GetSetting(APPNAME, App.Title, "CodTrans", "")
        .MONEDA = GetSetting(APPNAME, App.Title, "Moneda", "USD")
        .FormaCobroPago = GetSetting(APPNAME, App.Title, "Forma", "")
        .Responsable = GetSetting(APPNAME, App.Title, "Responsable", "")
        .AbrirArchivoenFormaDiferencial = GetSetting(APPNAME, App.Title, "AbrirArchivoenFormaDiferencial", False)
    End With
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Public Sub GuardaConfig()
    'Guarda Configuracion
    With gConfig
        SaveSetting APPNAME, App.Title, "Codcli", .CodCli
        SaveSetting APPNAME, App.Title, "CodTrans", .CodTrans
        SaveSetting APPNAME, App.Title, "Moneda", .MONEDA
        SaveSetting APPNAME, App.Title, "Forma", .FormaCobroPago
        SaveSetting APPNAME, App.Title, "Responsable", .Responsable
        SaveSetting APPNAME, App.Title, "AbrirArchivoenFormaDiferencial", .AbrirArchivoenFormaDiferencial
    End With
End Sub

'jeaa 26/05/2005
Public Sub CargarTareas( _
                ByVal grd As VSFlexGrid)
    Dim i As Integer

    With grd
        .Cols = 3
        .Rows = .FixedRows
        .AddItem .Rows & vbTab & "Tarea 1" & vbTab & "t1"
        .AddItem .Rows & vbTab & "Tarea 2" & vbTab & "t2"
        .FormatString = "^#|<Tarea|<Tabla"
        GNPoneNumFila grd, False
        AsignarTituloAColKey grd            'Para usar ColIndex
        AjustarAutoSize grd, -1, -1, 3000  'Ajusta automáticamente ancho de cols.
        'Oculta columnas innecesarias
        .ColHidden(.ColIndex("Tabla")) = True
    End With
End Sub

Public Sub CargaTipoTrans(ByRef Modulo As String, ByRef lst As ListBox)
    Dim rs As Recordset, Vector As Variant
    Dim numMod As Integer, i As Integer
    'Prepara la lista de tipos de transaccion
    lst.Clear
    Vector = Split(Modulo, ",")
    numMod = UBound(Vector, 1)
    If numMod = -1 Then
        Set rs = gobjMain.EmpresaActual.ListaGNTrans("", False, True)
        With rs
            If Not (.EOF) Then
                .MoveFirst
                Do Until .EOF
                    lst.AddItem !CodTrans & "  " & !NombreTrans
                    lst.ItemData(lst.NewIndex) = Len(!CodTrans)
                    .MoveNext
                Loop
            End If
        End With
        rs.Close
    Else
        For i = 0 To numMod
            Set rs = gobjMain.EmpresaActual.ListaGNTrans(CStr(Vector(i)), False, True)
            With rs
                If Not (.EOF) Then
                    .MoveFirst
                    Do Until .EOF
                        lst.AddItem !CodTrans & "  " & !NombreTrans
                        lst.ItemData(lst.NewIndex) = Len(!CodTrans)
                        .MoveNext
                    Loop
                End If
            End With
            rs.Close
        Next i
    End If
    Set rs = Nothing
End Sub

Public Sub EscribirIntervalosA()
    Dim file As String, NumFile As Long, i As Integer
    Dim LINEA As String
    On Error GoTo ErrTrap
    NumFile = FreeFile
    'file = IIf(Len(gConfigura.Archivo) > 0, gConfigura.Archivo, "C:\Archivos de programa\Sii4A\TablaComisiones.txt")
    file = IIf(Len(gConfigura.Archivo) > 0, gConfigura.Archivo, App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "TablaComisiones.txt")
    Close #NumFile
    'Proceso para escribir en un archivo de texto
    Open file For Output Access Write As #NumFile
        For i = 1 To 10
                LINEA = gComisiones(i).desde & ";" & gComisiones(i).hasta & ";" & _
                        gComisiones(i).Comision & ";" & gComisiones(i).ComisionC
                Print #NumFile, LINEA
        Next i
    Close NumFile
    LeerIntervalos 'Para actualizar los nuevos valores
    Exit Sub
    
ErrTrap:
    MsgBox Err.Description
    Exit Sub
End Sub

Public Sub LeerIntervalos()
    Dim file As String, NumFile As Long, i As Integer
    Dim LINEA As String, v As Variant, bandhaydatos As Boolean
    On Error GoTo ErrTrap
    
    NumFile = FreeFile
    'file = IIf(Len(gConfigura.Archivo) > 0, gConfigura.Archivo, "C:\Archivos de Programa\Sii4A\TablaComisiones")
    file = IIf(Len(gConfigura.Archivo) > 0, gConfigura.Archivo, App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "TablaComisiones.txt")
    gConfigura.Archivo = file
    'Proceso para leer un archivo primera y última línea casos especiales
    i = 0
    bandhaydatos = False
    Open file For Input As #NumFile
        Do While Not EOF(NumFile) And i < 10
            Line Input #NumFile, LINEA
            v = Split(LINEA, ";")
            i = i + 1
            gComisiones(i).desde = CCur("0" & v(0))
            gComisiones(i).hasta = CCur("0" & v(1))
            gComisiones(i).Comision = CCur("0" & v(2))
            gComisiones(i).ComisionC = CCur("0" & v(3))
            bandhaydatos = True
        Loop
    Close #NumFile

Continuar:
    'si i=0 significa que el archivo está vacío
    If Not (bandhaydatos) Then
        For i = 1 To 10
            gComisiones(i).desde = 0
            gComisiones(i).hasta = 0
            gComisiones(i).Comision = 0
            gComisiones(i).ComisionC = 0
        Next i
    End If
    Exit Sub
    
ErrTrap:
    'si no existe el archivo coloca ceros
    If Err.Number = 53 Then GoTo Continuar
    Exit Sub
End Sub

Public Sub GrabaConfig()
    With gConfigura
'        SaveSetting APPNAME, SECTION, "Archivo", .Archivo
'        SaveSetting APPNAME, SECTION, "ArchivoB", .ArchivoB
'        SaveSetting APPNAME, SECTION, "ArchivoC", .ArchivoC
'        SaveSetting APPNAME, SECTION, "ArchivoD", .ArchivoD
        
        
'        gobjMain.EmpresaActual.GNOpcion.AsignarValor "RutaTablaComisionesA", .Archivo
'        gobjMain.EmpresaActual.GNOpcion.AsignarValor "RutaTablaComisionesB", .ArchivoB
'        gobjMain.EmpresaActual.GNOpcion.AsignarValor "RutaTablaComisionesC", .ArchivoC
'        gobjMain.EmpresaActual.GNOpcion.AsignarValor "RutaTablaComisionesD", .ArchivoD

        
'        gobjMain.EmpresaActual.GNOpcion.Grabar
        
        
        
        '
        SaveSetting APPNAME, SECTION, "PorcenVendedorA", .PorcenVendedorA
        SaveSetting APPNAME, SECTION, "PorcenVendedorB", .PorcenVendedorB
        SaveSetting APPNAME, SECTION, "PorcenVendedorC", .PorcenVendedorC
        SaveSetting APPNAME, SECTION, "PorcenVendedorD", .PorcenVendedorD
        
        SaveSetting APPNAME, SECTION, "PorcenCobradorA", .PorcenCobradorA
        SaveSetting APPNAME, SECTION, "PorcenCobradorB", .PorcenCobradorB
        SaveSetting APPNAME, SECTION, "PorcenCobradorC", .PorcenCobradorC
        SaveSetting APPNAME, SECTION, "PorcenCobradorD", .PorcenCobradorD
    End With
End Sub

Public Sub RecuperarConfig()
    With gConfigura
        '.Archivo = GetSetting(APPNAME, SECTION, "Archivo", "C:\Archivos de Programa\Sii4A\TablaComisionesA.txt")
        '.ArchivoB = GetSetting(APPNAME, SECTION, "ArchivoB", "C:\Archivos de Programa\Sii4A\TablaComisionesB.txt")
        '.ArchivoC = GetSetting(APPNAME, SECTION, "ArchivoC", "C:\Archivos de Programa\Sii4A\TablaComisionesC.txt")
        '.ArchivoD = GetSetting(APPNAME, SECTION, "ArchivoD", "C:\Archivos de Programa\Sii4A\TablaComisionesD.txt")
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TablaComisionesA")) = 0 Then
            .Archivo = GetSetting(APPNAME, SECTION, "Archivo", "C:\Archivos de Programa\Sii4A\TablaComisionesA.txt")
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TablaComisionesB")) = 0 Then
            .ArchivoB = GetSetting(APPNAME, SECTION, "Archivo", "C:\Archivos de Programa\Sii4A\TablaComisionesB.txt")
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TablaComisionesC")) = 0 Then
            .ArchivoC = GetSetting(APPNAME, SECTION, "Archivo", "C:\Archivos de Programa\Sii4A\TablaComisionesC.txt")
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TablaComisionesD")) = 0 Then
            .ArchivoD = GetSetting(APPNAME, SECTION, "Archivo", "C:\Archivos de Programa\Sii4A\TablaComisionesD.txt")
        End If
        
        
'        .PorcenVendedorA = GetSetting(APPNAME, SECTION, "PorcenVendedorA", "0")
'        .PorcenVendedorB = GetSetting(APPNAME, SECTION, "PorcenVendedorB", "0")
'        .PorcenVendedorC = GetSetting(APPNAME, SECTION, "PorcenVendedorC", "0")
'        .PorcenVendedorD = GetSetting(APPNAME, SECTION, "PorcenVendedorD", "0")
'
'        .PorcenCobradorA = GetSetting(APPNAME, SECTION, "PORCENCobradorA", "0")
'        .PorcenCobradorB = GetSetting(APPNAME, SECTION, "PORCENCobradorB", "0")
'        .PorcenCobradorC = GetSetting(APPNAME, SECTION, "PORCENCobradorC", "0")
'        .PorcenCobradorD = GetSetting(APPNAME, SECTION, "PORCENCobradorD", "0")
        
    End With
End Sub

Public Function PreparaCadena(ByVal Cadena As String) As String
'Funcion que concatena apostrofes en una cadena separada por comas
Dim v As Variant, max As Integer, i As Integer
Dim Respuesta As String
    If Cadena = "" Then
        PreparaCadena = "''"
        Exit Function
    End If
    v = Split(Cadena, ",")
    max = UBound(v, 1)
    For i = 0 To max
        Respuesta = Respuesta & "'" & v(i) & "'" & ","
    Next i
    Respuesta = Left(Respuesta, Len(Respuesta) - 1) 'Quita la útima coma
    PreparaCadena = Respuesta
    
End Function

Public Sub RecuperarConfigIVAjusteAutomatico() '***jeaa 20/09/2006
    With gConfigIVAjusteAutomatico
        .CodTrans_AA = GetSetting(APPNAME, App.Title, "CodTrans_AA", "")
        .CodTrans_AAJ = GetSetting(APPNAME, App.Title, "CodTrans_AAJ", "")
        .CodTrans_ABJ = GetSetting(APPNAME, App.Title, "CodTrans_ABJ", "")
        .BandLineaAutoA = GetSetting(APPNAME, App.Title, "BandLineaAutoA", False)
        'jeaa 13/10/04
        .BandTotalizarItemA = GetSetting(APPNAME, App.Title, "BandTotalizarItemA", False)
    End With
End Sub

Public Sub GrabarConfigIVAjusteAutomatico() '*** jeaa 20/09/2006
    With gConfigIVAjusteAutomatico
        SaveSetting APPNAME, App.Title, "CodTrans_AA", .CodTrans_AA
        SaveSetting APPNAME, App.Title, "CodTrans_AAJ", .CodTrans_AAJ
        SaveSetting APPNAME, App.Title, "CodTrans_ABJ", .CodTrans_ABJ
        SaveSetting APPNAME, App.Title, "BandLineaAutoA", .BandLineaAutoA
        SaveSetting APPNAME, App.Title, "BandTotalizarItemA", .BandTotalizarItemA
    End With
End Sub


Public Sub CargaTransxTipoComprobante(ByRef lst As ListBox, ByVal Tipo As String)
    Dim rs As Recordset, sql As String
    Dim numMod As Integer, i As Integer
    'Prepara la lista de tipos de transaccion
    lst.Clear
        
            sql = "SELECT CodTrans, NombreTrans FROM  GnTrans WHERE  "
            sql = sql & " AnexoCodTipoComp =" & Tipo

            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                With rs
                If Not (.EOF) Then
                    .MoveFirst
                    Do Until .EOF
                        lst.AddItem !CodTrans & "  " & !NombreTrans
                        lst.ItemData(lst.NewIndex) = Len(!CodTrans)
                        .MoveNext
                    Loop
                End If
            End With
            rs.Close
    Set rs = Nothing
End Sub

Public Sub LeerIntervalosB()
    Dim file As String, NumFile As Long, i As Integer
    Dim LINEA As String, v As Variant, bandhaydatos As Boolean
    On Error GoTo ErrTrap
    
    NumFile = FreeFile
    'file = IIf(Len(gConfigura.ArchivoB) > 0, gConfigura.ArchivoB, "C:\Archivos de Programa\Sii4A\TablaComisionesB")
    file = IIf(Len(gConfigura.ArchivoB) > 0, gConfigura.ArchivoB, App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "TablaComisionesB.txt")
    gConfigura.ArchivoB = file
    'Proceso para leer un archivo primera y última línea casos especiales
    i = 0
    bandhaydatos = False
    Open file For Input As #NumFile
        Do While Not EOF(NumFile) And i < 10
            Line Input #NumFile, LINEA
            v = Split(LINEA, ";")
            i = i + 1
            gComisionesB(i).desde = CCur("0" & v(0))
            gComisionesB(i).hasta = CCur("0" & v(1))
            gComisionesB(i).Comision = CCur("0" & v(2))
            gComisionesB(i).ComisionC = CCur("0" & v(3))
            bandhaydatos = True
        Loop
    Close #NumFile

Continuar:
    'si i=0 significa que el archivo está vacío
    If Not (bandhaydatos) Then
        For i = 1 To 10
            gComisionesB(i).desde = 0
            gComisionesB(i).hasta = 0
            gComisionesB(i).Comision = 0
            gComisionesB(i).ComisionC = 0
        Next i
    End If
    Exit Sub
    
ErrTrap:
    'si no existe el archivo coloca ceros
    If Err.Number = 53 Then GoTo Continuar
    Exit Sub
End Sub

Public Sub EscribirIntervalosB()
    Dim file As String, NumFile As Long, i As Integer
    Dim LINEA As String
    On Error GoTo ErrTrap
    NumFile = FreeFile
'    file = IIf(Len(gConfigura.ArchivoB) > 0, gConfigura.ArchivoB, "C:\Archivos de programa\Sii4A\TablaComisionesB.txt")
    file = IIf(Len(gConfigura.ArchivoB) > 0, gConfigura.ArchivoB, App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "TablaComisionesB.txt")
    Close #NumFile
    'Proceso para escribir en un archivo de texto
    Open file For Output Access Write As #NumFile
        For i = 1 To 10
                LINEA = gComisionesB(i).desde & ";" & gComisionesB(i).hasta & ";" & _
                        gComisionesB(i).Comision & ";" & gComisionesB(i).ComisionC
                Print #NumFile, LINEA
        Next i
    Close NumFile
    LeerIntervalosB 'Para actualizar los nuevos valores
    Exit Sub
    
ErrTrap:
    MsgBox Err.Description
    Exit Sub
End Sub


Public Sub CargaTransRetencion(ByRef lst As ListBox, ByRef bandIVA As Boolean)
    Dim rs As Recordset, sql As String
    Dim numMod As Integer, i As Integer
    'Prepara la lista de tipos de transaccion
    lst.Clear
        
            sql = "SELECT CodRetencion, left(Descripcion,30) + ' Campo: '+ codf104 as des FROM  tsretencion "
            sql = sql & "WHERE bandIVA=" & IIf(bandIVA, 1, 0)
            sql = sql & " and len(codf104)>0 "
            sql = sql & " and bandValida=1  order by codsri "
           
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                With rs
                If Not (.EOF) Then
                    .MoveFirst
                    Do Until .EOF
                        lst.AddItem !CodRetencion & "  " & !Des
                        lst.ItemData(lst.NewIndex) = Len(!CodRetencion)
                        .MoveNext
                    Loop
                End If
            End With
            rs.Close
    Set rs = Nothing
End Sub

Public Sub EscribirIntervalosC()
    Dim file As String, NumFile As Long, i As Integer
    Dim LINEA As String
    On Error GoTo ErrTrap
    NumFile = FreeFile
    'file = IIf(Len(gConfigura.ArchivoC) > 0, gConfigura.ArchivoC, "C:\Archivos de programa\Sii4A\TablaComisionesC.txt")
    file = IIf(Len(gConfigura.ArchivoC) > 0, gConfigura.ArchivoC, App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "TablaComisionesC.txt")
    Close #NumFile
    'Proceso para escribir en un archivo de texto
    Open file For Output Access Write As #NumFile
        For i = 1 To 10
                LINEA = gComisionesC(i).desde & ";" & gComisionesC(i).hasta & ";" & _
                        gComisionesC(i).Comision & ";" & gComisionesC(i).ComisionC
                Print #NumFile, LINEA
        Next i
    Close NumFile
    LeerIntervalosC 'Para actualizar los nuevos valores
    Exit Sub
    
ErrTrap:
    MsgBox Err.Description
    Exit Sub
End Sub


Public Sub LeerIntervalosC()
    Dim file As String, NumFile As Long, i As Integer
    Dim LINEA As String, v As Variant, bandhaydatos As Boolean
    On Error GoTo ErrTrap
    
    NumFile = FreeFile
'    file = IIf(Len(gConfigura.Archivo) > 0, gConfigura.Archivo, "C:\Archivos de Programa\Sii4A\TablaComisionesC")
    file = IIf(Len(gConfigura.ArchivoC) > 0, gConfigura.ArchivoC, App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "TablaComisionesC.txt")
    gConfigura.ArchivoC = file
    'Proceso para leer un archivo primera y última línea casos especiales
    i = 0
    bandhaydatos = False
    Open file For Input As #NumFile
        Do While Not EOF(NumFile) And i < 10
            Line Input #NumFile, LINEA
            v = Split(LINEA, ";")
            i = i + 1
            gComisionesC(i).desde = CCur("0" & v(0))
            gComisionesC(i).hasta = CCur("0" & v(1))
            gComisionesC(i).Comision = CCur("0" & v(2))
            gComisionesC(i).ComisionC = CCur("0" & v(3))
            bandhaydatos = True
        Loop
    Close #NumFile

Continuar:
    'si i=0 significa que el archivo está vacío
    If Not (bandhaydatos) Then
        For i = 1 To 10
            gComisionesC(i).desde = 0
            gComisionesC(i).hasta = 0
            gComisionesC(i).Comision = 0
            gComisionesC(i).ComisionC = 0
        Next i
    End If
    Exit Sub
    
ErrTrap:
    'si no existe el archivo coloca ceros
    If Err.Number = 53 Then GoTo Continuar
    Exit Sub
End Sub

Public Sub EscribirIntervalosD()
    Dim file As String, NumFile As Long, i As Integer
    Dim LINEA As String
    On Error GoTo ErrTrap
    NumFile = FreeFile
    'file = IIf(Len(gConfigura.ArchivoD) > 0, gConfigura.ArchivoD, "C:\Archivos de programa\Sii4A\TablaComisionesD.txt")
    file = IIf(Len(gConfigura.ArchivoD) > 0, gConfigura.ArchivoD, App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "TablaComisionesD.txt")
    Close #NumFile
    'Proceso para escribir en un archivo de texto
    Open file For Output Access Write As #NumFile
        For i = 1 To 10
                LINEA = gComisionesD(i).desde & ";" & gComisionesD(i).hasta & ";" & _
                        gComisionesD(i).Comision & ";" & gComisionesD(i).ComisionC
                Print #NumFile, LINEA
        Next i
    Close NumFile
    LeerIntervalosD 'Para actualizar los nuevos valores
    Exit Sub
    
ErrTrap:
    MsgBox Err.Description
    Exit Sub
End Sub


Public Sub LeerIntervalosD()
    Dim file As String, NumFile As Long, i As Integer
    Dim LINEA As String, v As Variant, bandhaydatos As Boolean
    On Error GoTo ErrTrap
    
    NumFile = FreeFile
'    file = IIf(Len(gConfigura.Archivo) > 0, gConfigura.Archivo, "C:\Archivos de Programa\Sii4A\TablaComisionesD")
    file = IIf(Len(gConfigura.ArchivoD) > 0, gConfigura.ArchivoD, App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "TablaComisionesD.txt")
    gConfigura.ArchivoD = file
    'Proceso para leer un archivo primera y última línea casos especiales
    i = 0
    bandhaydatos = False
    Open file For Input As #NumFile
        Do While Not EOF(NumFile) And i < 10
            Line Input #NumFile, LINEA
            v = Split(LINEA, ";")
            i = i + 1
            gComisionesD(i).desde = CCur("0" & v(0))
            gComisionesD(i).hasta = CCur("0" & v(1))
            gComisionesD(i).Comision = CCur("0" & v(2))
            gComisionesD(i).ComisionC = CCur("0" & v(3))
            bandhaydatos = True
        Loop
    Close #NumFile

Continuar:
    'si i=0 significa que el archivo está vacío
    If Not (bandhaydatos) Then
        For i = 1 To 10
            gComisionesD(i).desde = 0
            gComisionesD(i).hasta = 0
            gComisionesD(i).Comision = 0
            gComisionesD(i).ComisionC = 0
        Next i
    End If
    Exit Sub
    
ErrTrap:
    'si no existe el archivo coloca ceros
    If Err.Number = 53 Then GoTo Continuar
    Exit Sub
End Sub

Public Sub CargaTransxTipoTrans(ByRef lst As ListBox, ByVal Tipo As String, Optional ByVal Tipo1 As String)
    Dim rs As Recordset, sql As String
    Dim numMod As Integer, i As Integer
    'Prepara la lista de tipos de transaccion
    lst.Clear

            sql = "SELECT CodTrans, NombreTrans FROM  GnTrans WHERE  "
            sql = sql & " AnexoCodTipoTrans =" & Tipo
            If Len(Tipo1) > 0 Then
                sql = sql & " OR AnexoCodTipoTrans =" & Tipo1
            End If

            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                With rs
                If Not (.EOF) Then
                    .MoveFirst
                    Do Until .EOF
                        lst.AddItem !CodTrans & "  " & !NombreTrans
                        lst.ItemData(lst.NewIndex) = Len(!CodTrans)
                        .MoveNext
                    Loop
                End If
            End With
            rs.Close
    Set rs = Nothing
End Sub

Public Sub CargaRetencion(ByRef lst As ListBox)
    Dim rs As Recordset
    lst.Clear
    Set rs = gobjMain.EmpresaActual.ListaTSRetencion(True, True)
    With rs
        If Not (.EOF) Then
            .MoveFirst
            Do Until .EOF
               If Left(!CodRetencion, 2) = "IV" Then
                    lst.AddItem !CodRetencion & "  " & !Descripcion
                    lst.ItemData(lst.NewIndex) = Len(!CodRetencion)
               End If
               .MoveNext
           Loop
        End If
    End With
    rs.Close
End Sub

'''''Public Sub LeerIntervalosJefeA()
'''''    Dim file As String, NumFile As Long, i As Integer
'''''    Dim linea As String, v As Variant, bandhaydatos As Boolean
'''''    On Error GoTo ErrTrap
'''''
'''''    NumFile = FreeFile
'''''    'file = IIf(Len(gConfigura.ArchivoB) > 0, gConfigura.ArchivoB, "C:\Archivos de Programa\Sii4A\TablaComisionesB")
'''''    file = IIf(Len(gConfigura.ArchivoJefeA) > 0, gConfigura.ArchivoJefeA, App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "TablaComisionesJefeA.txt")
'''''    gConfigura.ArchivoJefeA = file
'''''    'Proceso para leer un archivo primera y última línea casos especiales
'''''    i = 0
'''''    bandhaydatos = False
'''''    Open file For Input As #NumFile
'''''        Do While Not EOF(NumFile) And i < 10
'''''            Line Input #NumFile, linea
'''''            v = Split(linea, ";")
'''''            i = i + 1
'''''            gComisionesJefeA(i).desde = CCur("0" & v(0))
'''''            gComisionesJefeA(i).hasta = CCur("0" & v(1))
'''''            gComisionesJefeA(i).Comision = CCur("0" & v(2))
'''''            bandhaydatos = True
'''''        Loop
'''''    Close #NumFile
'''''
'''''Continuar:
'''''    'si i=0 significa que el archivo está vacío
'''''    If Not (bandhaydatos) Then
'''''        For i = 1 To 10
'''''            gComisionesJefeA(i).desde = 0
'''''            gComisionesJefeA(i).hasta = 0
'''''            gComisionesJefeA(i).Comision = 0
'''''        Next i
'''''    End If
'''''    Exit Sub
'''''
'''''ErrTrap:
'''''    'si no existe el archivo coloca ceros
'''''    If Err.Number = 53 Then GoTo Continuar
'''''    Exit Sub
'''''End Sub
'''''
'''''
'''''Public Sub LeerIntervalosJefeB()
'''''    Dim file As String, NumFile As Long, i As Integer
'''''    Dim linea As String, v As Variant, bandhaydatos As Boolean
'''''    On Error GoTo ErrTrap
'''''
'''''    NumFile = FreeFile
'''''    'file = IIf(Len(gConfigura.ArchivoB) > 0, gConfigura.ArchivoB, "C:\Archivos de Programa\Sii4A\TablaComisionesB")
'''''    file = IIf(Len(gConfigura.ArchivoJefeB) > 0, gConfigura.ArchivoJefeB, App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "TablaComisionesJefeB.txt")
'''''    gConfigura.ArchivoJefeB = file
'''''    'Proceso para leer un archivo primera y última línea casos especiales
'''''    i = 0
'''''    bandhaydatos = False
'''''    Open file For Input As #NumFile
'''''        Do While Not EOF(NumFile) And i < 10
'''''            Line Input #NumFile, linea
'''''            v = Split(linea, ";")
'''''            i = i + 1
'''''            gComisionesJefeB(i).desde = CCur("0" & v(0))
'''''            gComisionesJefeB(i).hasta = CCur("0" & v(1))
'''''            gComisionesJefeB(i).Comision = CCur("0" & v(2))
'''''            bandhaydatos = True
'''''        Loop
'''''    Close #NumFile
'''''
'''''Continuar:
'''''    'si i=0 significa que el archivo está vacío
'''''    If Not (bandhaydatos) Then
'''''        For i = 1 To 10
'''''            gComisionesJefeB(i).desde = 0
'''''            gComisionesJefeB(i).hasta = 0
'''''            gComisionesJefeB(i).Comision = 0
'''''        Next i
'''''    End If
'''''    Exit Sub
'''''
'''''ErrTrap:
'''''    'si no existe el archivo coloca ceros
'''''    If Err.Number = 53 Then GoTo Continuar
'''''    Exit Sub
'''''End Sub
'''''
'''''Public Sub LeerIntervalosJefeC()
'''''    Dim file As String, NumFile As Long, i As Integer
'''''    Dim linea As String, v As Variant, bandhaydatos As Boolean
'''''    On Error GoTo ErrTrap
'''''
'''''    NumFile = FreeFile
'''''    'file = IIf(Len(gConfigura.ArchivoB) > 0, gConfigura.ArchivoB, "C:\Archivos de Programa\Sii4A\TablaComisionesB")
'''''    file = IIf(Len(gConfigura.ArchivoJefeC) > 0, gConfigura.ArchivoJefeC, App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "TablaComisionesJefeC.txt")
'''''    gConfigura.ArchivoJefeC = file
'''''    'Proceso para leer un archivo primera y última línea casos especiales
'''''    i = 0
'''''    bandhaydatos = False
'''''    Open file For Input As #NumFile
'''''        Do While Not EOF(NumFile) And i < 10
'''''            Line Input #NumFile, linea
'''''            v = Split(linea, ";")
'''''            i = i + 1
'''''            gComisionesJefeC(i).desde = CCur("0" & v(0))
'''''            gComisionesJefeC(i).hasta = CCur("0" & v(1))
'''''            gComisionesJefeC(i).Comision = CCur("0" & v(2))
'''''            bandhaydatos = True
'''''        Loop
'''''    Close #NumFile
'''''
'''''Continuar:
'''''    'si i=0 significa que el archivo está vacío
'''''    If Not (bandhaydatos) Then
'''''        For i = 1 To 10
'''''            gComisionesJefeC(i).desde = 0
'''''            gComisionesJefeC(i).hasta = 0
'''''            gComisionesJefeC(i).Comision = 0
'''''        Next i
'''''    End If
'''''    Exit Sub
'''''
'''''ErrTrap:
'''''    'si no existe el archivo coloca ceros
'''''    If Err.Number = 53 Then GoTo Continuar
'''''    Exit Sub
'''''End Sub
'''''
'''''Public Sub LeerIntervalosJefeD()
'''''    Dim file As String, NumFile As Long, i As Integer
'''''    Dim linea As String, v As Variant, bandhaydatos As Boolean
'''''    On Error GoTo ErrTrap
'''''
'''''    NumFile = FreeFile
'''''    'file = IIf(Len(gConfigura.ArchivoB) > 0, gConfigura.ArchivoB, "C:\Archivos de Programa\Sii4A\TablaComisionesB")
'''''    file = IIf(Len(gConfigura.ArchivoJefeD) > 0, gConfigura.ArchivoJefeD, App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "TablaComisionesJefeD.txt")
'''''    gConfigura.ArchivoJefeD = file
'''''    'Proceso para leer un archivo primera y última línea casos especiales
'''''    i = 0
'''''    bandhaydatos = False
'''''    Open file For Input As #NumFile
'''''        Do While Not EOF(NumFile) And i < 10
'''''            Line Input #NumFile, linea
'''''            v = Split(linea, ";")
'''''            i = i + 1
'''''            gComisionesJefeD(i).desde = CCur("0" & v(0))
'''''            gComisionesJefeD(i).hasta = CCur("0" & v(1))
'''''            gComisionesJefeD(i).Comision = CCur("0" & v(2))
'''''            bandhaydatos = True
'''''        Loop
'''''    Close #NumFile
'''''
'''''Continuar:
'''''    'si i=0 significa que el archivo está vacío
'''''    If Not (bandhaydatos) Then
'''''        For i = 1 To 10
'''''            gComisionesJefeD(i).desde = 0
'''''            gComisionesJefeD(i).hasta = 0
'''''            gComisionesJefeD(i).Comision = 0
'''''        Next i
'''''    End If
'''''    Exit Sub
'''''
'''''ErrTrap:
'''''    'si no existe el archivo coloca ceros
'''''    If Err.Number = 53 Then GoTo Continuar
'''''    Exit Sub
'''''End Sub



Public Sub LeerIntervalosJefe(ByVal TipoTabla As String)
    Dim file As String, NumFile As Long, i As Integer
    Dim LINEA As String, v As Variant, bandhaydatos As Boolean
    Dim j As Integer
    On Error GoTo ErrTrap
    
    NumFile = FreeFile
    'file = IIf(Len(gConfigura.ArchivoB) > 0, gConfigura.ArchivoB, "C:\Archivos de Programa\Sii4A\TablaComisionesB")
    Select Case TipoTabla
    Case "A"
        file = IIf(Len(gConfiguraJefe.ArchivoJefeA) > 0, gConfiguraJefe.ArchivoJefeA, App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "TablaComisionesJefeA.txt")
        gConfiguraJefe.ArchivoJefeA = file
    Case "B"
        file = IIf(Len(gConfiguraJefe.ArchivoJefeB) > 0, gConfiguraJefe.ArchivoJefeB, App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "TablaComisionesJefeB.txt")
        gConfiguraJefe.ArchivoJefeB = file
    Case "C"
        file = IIf(Len(gConfiguraJefe.ArchivoJefeC) > 0, gConfiguraJefe.ArchivoJefeC, App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "TablaComisionesJefeC.txt")
        gConfiguraJefe.ArchivoJefeC = file
    Case "D"
        file = IIf(Len(gConfiguraJefe.ArchivoJefeD) > 0, gConfiguraJefe.ArchivoJefeD, App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "TablaComisionesJefeD.txt")
        gConfiguraJefe.ArchivoJefeD = file
    End Select
    
    'Proceso para leer un archivo primera y última línea casos especiales
    i = 0
    bandhaydatos = False
    Open file For Input As #NumFile
        Do While Not EOF(NumFile) And i < 10
            Line Input #NumFile, LINEA
            v = Split(LINEA, ";")
            i = i + 1
            gComisionesJefe(i).desde = CCur("0" & v(0))
            gComisionesJefe(i).hasta = CCur("0" & v(1))
            gComisionesJefe(i).Comision = CCur("0" & v(2))
            
            bandhaydatos = True
        Loop
    Close #NumFile
    For j = 1 To 10
        Select Case TipoTabla
            Case "A": gComisionesJefeA(j) = gComisionesJefe(j)
            Case "B": gComisionesJefeB(j) = gComisionesJefe(j)
            Case "C": gComisionesJefeC(j) = gComisionesJefe(j)
            Case "D": gComisionesJefeD(j) = gComisionesJefe(j)
        End Select
    Next j
Continuar:
    'si i=0 significa que el archivo está vacío
    If Not (bandhaydatos) Then
        For i = 1 To 10
            gComisionesJefe(i).desde = 0
            gComisionesJefe(i).hasta = 0
            gComisionesJefe(i).Comision = 0
        Next i
    End If
    For j = 1 To 10
        Select Case TipoTabla
            Case "A": gComisionesJefeA(j) = gComisionesJefe(j)
            Case "B": gComisionesJefeB(j) = gComisionesJefe(j)
            Case "C": gComisionesJefeC(j) = gComisionesJefe(j)
            Case "D": gComisionesJefeD(j) = gComisionesJefe(j)
        End Select
    Next j
    Exit Sub
    
ErrTrap:
    'si no existe el archivo coloca ceros
    If Err.Number = 53 Then GoTo Continuar
    Exit Sub
End Sub

Public Sub EscribirIntervalosJefe(ByVal TipoTabla As String)
    Dim file As String, NumFile As Long, i As Integer
    Dim LINEA As String, v As Variant, bandhaydatos As Boolean
    Dim j As Integer
    On Error GoTo ErrTrap
    
    NumFile = FreeFile
    'file = IIf(Len(gConfigura.Archivo) > 0, gConfigura.Archivo, "C:\Archivos de Programa\Sii4A\TablaComisiones")
    Select Case TipoTabla
    Case "A"
        file = IIf(Len(gConfiguraJefe.ArchivoJefeA) > 0, gConfiguraJefe.ArchivoJefeA, App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "TablaComisionesJefeA.txt")
        gConfiguraJefe.ArchivoJefeA = file
    Case "B"
        file = IIf(Len(gConfiguraJefe.ArchivoJefeB) > 0, gConfiguraJefe.ArchivoJefeB, App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "TablaComisionesJefeB.txt")
        gConfiguraJefe.ArchivoJefeB = file
    Case "C"
        file = IIf(Len(gConfiguraJefe.ArchivoJefeC) > 0, gConfiguraJefe.ArchivoJefeC, App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "TablaComisionesJefeC.txt")
        gConfiguraJefe.ArchivoJefeC = file
    Case "D"
        file = IIf(Len(gConfiguraJefe.ArchivoJefeD) > 0, gConfiguraJefe.ArchivoJefeD, App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "TablaComisionesJefeD.txt")
        gConfiguraJefe.ArchivoJefeD = file
    End Select
    Close #NumFile
    'Proceso para escribir en un archivo de texto
    Open file For Output Access Write As #NumFile
        For i = 1 To 10
                Select Case TipoTabla
                    Case "A":   LINEA = gComisionesJefeA(i).desde & ";" & gComisionesJefeA(i).hasta & ";" & gComisionesJefeA(i).Comision
                    Case "B":   LINEA = gComisionesJefeB(i).desde & ";" & gComisionesJefeB(i).hasta & ";" & gComisionesJefeB(i).Comision
                    Case "C":   LINEA = gComisionesJefeC(i).desde & ";" & gComisionesJefeC(i).hasta & ";" & gComisionesJefeC(i).Comision
                    Case "D":   LINEA = gComisionesJefeD(i).desde & ";" & gComisionesJefeD(i).hasta & ";" & gComisionesJefeD(i).Comision
                End Select
                Print #NumFile, LINEA
        Next i
    Close NumFile
    LeerIntervalosJefe TipoTabla
    Exit Sub
ErrTrap:
    MsgBox Err.Description
    Exit Sub
End Sub


Public Sub CargaTipoTransxTipoTrans(ByRef Modulo As String, ByRef lst As ListBox, tipoTrans As String)
    Dim rs As Recordset, Vector As Variant
    Dim numMod As Integer, i As Integer
    'Prepara la lista de tipos de transaccion
    lst.Clear
    Vector = Split(Modulo, ",")
    numMod = UBound(Vector, 1)
    If numMod = -1 Then
        Set rs = gobjMain.EmpresaActual.ListaGNTrans("", False, True)
        With rs
            If Not (.EOF) Then
                .MoveFirst
                Do Until .EOF
                    lst.AddItem !CodTrans & "  " & !NombreTrans
                    lst.ItemData(lst.NewIndex) = Len(!CodTrans)
                    .MoveNext
                Loop
            End If
        End With
        rs.Close
    Else
        For i = 0 To numMod
            Set rs = gobjMain.EmpresaActual.ListaGNTransTipoTRans(CStr(Vector(i)), False, True, tipoTrans)
            With rs
                If Not (.EOF) Then
                    .MoveFirst
                    Do Until .EOF
                        lst.AddItem !CodTrans & "  " & !NombreTrans
                        lst.ItemData(lst.NewIndex) = Len(!CodTrans)
                        .MoveNext
                    Loop
                End If
            End With
            rs.Close
        Next i
    End If
    Set rs = Nothing
End Sub


Public Function ListaTramitesSRI() As Variant
    Dim v As Variant
    Dim fila As Integer
    Dim cad As String
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("RealizarReporteRangos")) > 0 Then
        If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("RealizarReporteRangos") = 0 Then
            If Date > gobjMain.EmpresaActual.GNOpcion.FechaCaducidad_AutoImp Then
                cad = "D"
            Else
                cad = "F"
            End If
        Else
            If Date > gobjMain.EmpresaActual.GNOpcion.FechaCaducidad_AutoImp Then
                cad = "D"
            Else
                cad = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TramitesPosiblesSRI")
            End If
        End If
    Else
        cad = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TramitesPosiblesSRI")
    End If
    
    
     
     'If Len(cad) = 0 Then cad = "B,C,D,E,F"
     ReDim v(2, 0)
    fila = 0
    
    If InStr(1, cad, "A") Then
        v(0, fila) = "6"
        v(1, fila) = "Solicitud de Autorización"
        v(2, fila) = "A"
        fila = fila + 1
        ReDim Preserve v(2, fila)
    End If
    
    If InStr(1, cad, "B") Then
        v(0, fila) = "7"
        v(1, fila) = "Solicitud de Autorización por Cambio de Software "
        v(2, fila) = "B"
        fila = fila + 1
        ReDim Preserve v(2, fila)
    End If
    
    If InStr(1, cad, "C") Then
        If DatePart("m", Date) = DatePart("m", gobjMain.EmpresaActual.GNOpcion.FechaAutorizacion_AutoImp) Then
            v(0, fila) = "8"
            v(1, fila) = "Renovación de Autorización"
            v(2, fila) = "C"
            fila = fila + 1
            ReDim Preserve v(2, fila)
        Else
        
        End If
    End If
    
    If InStr(1, cad, "D") Then
        v(0, fila) = "9"
        v(1, fila) = "Baja de Autorización"
        v(2, fila) = "D"
        fila = fila + 1
        ReDim Preserve v(2, fila)
    End If
    
    If InStr(1, cad, "E") Then
        v(0, fila) = "10"
        v(1, fila) = "Inclusión de Puntos y/o Documentos"
        v(2, fila) = "E"
        fila = fila + 1
        ReDim Preserve v(2, fila)
    End If

    If InStr(1, cad, "F") Then
        v(0, fila) = "11"
        v(1, fila) = "Eliminación de Puntos y/o Documentos"
        v(2, fila) = "F"
        fila = fila + 1
        ReDim Preserve v(2, fila)
    End If
    If fila > 0 Then
        ReDim Preserve v(2, fila - 1)
    End If
    
    
   ListaTramitesSRI = v
End Function

     
 Public Sub MiGetRowsRep(ByVal rs As Recordset, grd As VSFlexGrid)
    grd.LoadArray MiGetRows(rs)
    grd.Redraw = True
    ConfigTipoDatoCol grd, rs
End Sub


Public Sub ConfigTipoDatoCol(grd As VSFlexGrid, rs As Recordset)
    Dim f As Field, i As Integer, Tipo As Integer, movio As Boolean
    If Not (rs.EOF And rs.BOF) Then 'Si no esta vacio
        
        rs.MoveFirst
        For i = 0 To rs.Fields.count - 1
            'Mueve hasta una fila que no tenga Null en éste campo i.
            movio = False
            Do While IsNull(rs.Fields(i).value)
                rs.MoveNext
                movio = True
                If rs.EOF Then Exit Do
            Loop
            
            If rs.EOF Then
                Tipo = flexDTString
            Else
                 Select Case VarType(rs.Fields(i).value)
                 Case vbInteger
                     Tipo = flexDTShort
                 Case vbLong
                     Tipo = flexDTLong
                 Case vbSingle
                     Tipo = flexDTSingle
                 Case vbDouble
                     Tipo = flexDTDouble
                 Case vbCurrency
                     Tipo = flexDTCurrency
                 Case vbBoolean
                     Tipo = flexDTBoolean
                 Case vbDate
                     Tipo = flexDTDate
                 Case Else   'Tipo string
                     Tipo = flexDTString
                End Select
            End If
            grd.ColDataType(i + 1) = Tipo  ' Primera fila es la numeracion
            
            If movio Or rs.EOF Then rs.MoveFirst  'Regresa al primer registro para siguiente columna
        Next i
        Set f = Nothing
    End If
End Sub

Public Function ImprimeTrans(ByVal gc As GNComprobante, ByRef objImp As Object, ByVal Plantilla As String, ByVal msg As String, ByVal idAsignado As String, ByVal sec As Integer) As Boolean
   Dim crear As Boolean
    On Error GoTo ErrTrap
    'Si no tiene TransID quere decir que no está grabada
    If (gc.TransID = 0) Or gc.Modificado Then
        MsgBox MSGERR_NOGRABADO, vbInformation
        ImprimeTrans = False
        Exit Function
    End If
    'Solo por primera vez o cuando cambia la librería de impresión
    '  crea una instancia del objeto para la impresión
    crear = (objImp Is Nothing)
    If Not crear Then crear = (objImp.NombreDLL <> "gnprintg.dll")
    If crear Then
        Set objImp = Nothing
        Set objImp = CreateObject("Gnprintg.PrintTrans")
    End If
    MensajeStatus MSG_PREPARA, vbHourglass
    objImp.PrintNotificacionNew gobjMain.EmpresaActual, True, sec, idAsignado, "", 0, gc, Plantilla, msg


    MensajeStatus
    ImprimeTrans = True
    Exit Function
ErrTrap:
    MensajeStatus
    Select Case Err.Number
    Case ERR_NOIMPRIME, ERR_NOIMPRIME2, ERR_NOIMPRIME3, ERR_NOHAYCODIGO
        DispErr
    Case Else
        MsgBox MSGERR_NOIMPRIME2, vbInformation
    End Select
    ImprimeTrans = False
    Exit Function
End Function



Public Sub EscribirIntervalosGnOpcion(NumTabla As String)
    Dim file As String, NumFile As Long, i As Integer
    Dim LINEA As String
    On Error GoTo ErrTrap
    LINEA = ""
    Select Case NumTabla
     Case "A"
        For i = 1 To 10
                LINEA = LINEA & gComisiones(i).desde & "," & gComisiones(i).hasta & "," & gComisiones(i).Comision & "," & gComisiones(i).ComisionC & "," & gComisiones(i).ComisionSC & ";"
        Next i
     Case "B"
        For i = 1 To 10
                LINEA = LINEA & gComisionesB(i).desde & "," & gComisionesB(i).hasta & "," & gComisionesB(i).Comision & "," & gComisionesB(i).ComisionC & "," & gComisionesB(i).ComisionSC & ";"
        Next i
     Case "C"
        For i = 1 To 10
                LINEA = LINEA & gComisionesC(i).desde & "," & gComisionesC(i).hasta & "," & gComisionesC(i).Comision & "," & gComisionesC(i).ComisionC & "," & gComisionesC(i).ComisionSC & ";"
        Next i
     Case "D"
        For i = 1 To 10
                LINEA = LINEA & gComisionesD(i).desde & "," & gComisionesD(i).hasta & "," & gComisionesD(i).Comision & "," & gComisionesD(i).ComisionC & "," & gComisionesD(i).ComisionSC & ";"
        Next i
    End Select
     gobjMain.EmpresaActual.GNOpcion.AsignarValor "TablaComisiones" & NumTabla, LINEA
     gobjMain.EmpresaActual.GNOpcion.GrabarGNOpcion2
    'Close NumFile
    LeerIntervalosGnOpcion NumTabla
    Exit Sub
    
ErrTrap:
    MsgBox Err.Description
    Exit Sub
End Sub


Public Sub LeerIntervalosGnOpcion(NumTabla As String)
    Dim cad As String, NumFile As Long, i As Integer, W As Variant
    Dim LINEA As String, v As Variant, bandhaydatos As Boolean
    On Error GoTo ErrTrap
    
    
    cad = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TablaComisiones" & NumTabla)
    Select Case NumTabla
    Case "A"
        If Len(cad) > 0 Then
            bandhaydatos = False
            W = Split(cad, ";")
            For i = 1 To UBound(W)
                    v = Split(W(i - 1), ",")
                    gComisiones(i).desde = CCur("0" & v(0))
                    gComisiones(i).hasta = CCur("0" & v(1))
                    gComisiones(i).Comision = CCur("0" & v(2))
                    gComisiones(i).ComisionC = CCur("0" & v(3))
                    gComisiones(i).ComisionSC = CCur("0" & v(4))
                    bandhaydatos = True
            Next i
        Else
            If Not (bandhaydatos) Then
                For i = 1 To 10
                    gComisiones(i).desde = 0
                    gComisiones(i).hasta = 0
                    gComisiones(i).Comision = 0
                    gComisiones(i).ComisionC = 0
                    gComisiones(i).ComisionSC = 0
                Next i
            End If
        End If
    Case "B"
        If Len(cad) > 0 Then
            bandhaydatos = False
            W = Split(cad, ";")
            For i = 1 To UBound(W)
                    v = Split(W(i - 1), ",")
                    gComisionesB(i).desde = CCur("0" & v(0))
                    gComisionesB(i).hasta = CCur("0" & v(1))
                    gComisionesB(i).Comision = CCur("0" & v(2))
                    gComisionesB(i).ComisionC = CCur("0" & v(3))
                    gComisionesB(i).ComisionSC = CCur("0" & v(4))
                    bandhaydatos = True
            Next i
        Else
            If Not (bandhaydatos) Then
                For i = 1 To 10
                    gComisionesB(i).desde = 0
                    gComisionesB(i).hasta = 0
                    gComisionesB(i).Comision = 0
                    gComisionesB(i).ComisionC = 0
                    gComisionesB(i).ComisionSC = 0
                Next i
            End If
        End If
    Case "C"
        If Len(cad) > 0 Then
            bandhaydatos = False
            W = Split(cad, ";")
            For i = 1 To UBound(W)
                    v = Split(W(i - 1), ",")
                    gComisionesC(i).desde = CCur("0" & v(0))
                    gComisionesC(i).hasta = CCur("0" & v(1))
                    gComisionesC(i).Comision = CCur("0" & v(2))
                    gComisionesC(i).ComisionC = CCur("0" & v(3))
                    gComisionesC(i).ComisionSC = CCur("0" & v(4))
                    bandhaydatos = True
            Next i
        Else
            If Not (bandhaydatos) Then
                For i = 1 To 10
                    gComisionesC(i).desde = 0
                    gComisionesC(i).hasta = 0
                    gComisionesC(i).Comision = 0
                    gComisionesC(i).ComisionC = 0
                    gComisionesC(i).ComisionSC = 0
                Next i
            End If
        End If
    Case "D"
        If Len(cad) > 0 Then
            bandhaydatos = False
            W = Split(cad, ";")
            For i = 1 To UBound(W)
                    v = Split(W(i - 1), ",")
                    gComisionesD(i).desde = CCur("0" & v(0))
                    gComisionesD(i).hasta = CCur("0" & v(1))
                    gComisionesD(i).Comision = CCur("0" & v(2))
                    gComisionesD(i).ComisionC = CCur("0" & v(3))
                    gComisionesD(i).ComisionSC = CCur("0" & v(4))
                    bandhaydatos = True
            Next i
        Else
            If Not (bandhaydatos) Then
                For i = 1 To 10
                    gComisionesD(i).desde = 0
                    gComisionesD(i).hasta = 0
                    gComisionesD(i).Comision = 0
                    gComisionesD(i).ComisionC = 0
                    gComisionesD(i).ComisionSC = 0
                Next i
            End If
        End If
    End Select
    Exit Sub
    
ErrTrap:
    'si no existe el archivo coloca ceros

    Exit Sub
End Sub

Public Sub EscribirIntervalosJefeGnOpcion(NumTabla As String)
    Dim file As String, NumFile As Long, i As Integer
    Dim LINEA As String
    On Error GoTo ErrTrap
    LINEA = ""
    Select Case NumTabla
     Case "A"
        For i = 1 To 10
                LINEA = LINEA & gComisionesJefe(i).desde & "," & gComisionesJefe(i).hasta & "," & gComisionesJefe(i).Comision & "," & gComisionesJefe(i).ComisionC & ";"
        Next i
     Case "B"
        For i = 1 To 10
                LINEA = LINEA & gComisionesJefeB(i).desde & "," & gComisionesJefeB(i).hasta & "," & gComisionesJefeB(i).Comision & "," & gComisionesJefeB(i).ComisionC & ";"
        Next i
     Case "C"
        For i = 1 To 10
                LINEA = LINEA & gComisionesJefeC(i).desde & "," & gComisionesJefeC(i).hasta & "," & gComisionesJefeC(i).Comision & "," & gComisionesJefeC(i).ComisionC & ";"
        Next i
     Case "D"
        For i = 1 To 10
                LINEA = LINEA & gComisionesJefeD(i).desde & "," & gComisionesJefeD(i).hasta & "," & gComisionesJefeD(i).Comision & "," & gComisionesJefeD(i).ComisionC & ";"
        Next i
    End Select
     gobjMain.EmpresaActual.GNOpcion.AsignarValor "TablaComisionesJefe" & NumTabla, LINEA
     gobjMain.EmpresaActual.GNOpcion.GrabarGNOpcion2
    'Close NumFile
    LeerIntervalosGnOpcion NumTabla
    Exit Sub
    
ErrTrap:
    MsgBox Err.Description
    Exit Sub
End Sub

Public Sub LeerIntervaloJefesGnOpcion(NumTabla As String)
    Dim cad As String, NumFile As Long, i As Integer, W As Variant
    Dim LINEA As String, v As Variant, bandhaydatos As Boolean
    On Error GoTo ErrTrap
    
    
    cad = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TablaComisionesJefe" & NumTabla)
    Select Case NumTabla
    Case "A"
        If Len(cad) > 0 Then
            bandhaydatos = False
            W = Split(cad, ";")
            For i = 1 To UBound(W)
                    v = Split(W(i - 1), ",")
                    gComisionesJefe(i).desde = CCur("0" & v(0))
                    gComisionesJefe(i).hasta = CCur("0" & v(1))
                    gComisionesJefe(i).Comision = CCur("0" & v(2))
                    gComisionesJefe(i).ComisionC = CCur("0" & v(3))
                    gComisionesJefe(i).ComisionSC = CCur("0" & v(4))
                    bandhaydatos = True
            Next i
        Else
            If Not (bandhaydatos) Then
                For i = 1 To 10
                    gComisionesJefe(i).desde = 0
                    gComisionesJefe(i).hasta = 0
                    gComisionesJefe(i).Comision = 0
                    gComisionesJefe(i).ComisionC = 0
                    gComisionesJefe(i).ComisionSC = 0
                Next i
            End If
        End If
    Case "B"
        If Len(cad) > 0 Then
            bandhaydatos = False
            W = Split(cad, ";")
            For i = 1 To UBound(W)
                    v = Split(W(i - 1), ",")
                    gComisionesJefeB(i).desde = CCur("0" & v(0))
                    gComisionesJefeB(i).hasta = CCur("0" & v(1))
                    gComisionesJefeB(i).Comision = CCur("0" & v(2))
                    gComisionesJefeB(i).ComisionC = CCur("0" & v(3))
                    gComisionesJefeB(i).ComisionSC = CCur("0" & v(4))
                    bandhaydatos = True
            Next i
        Else
            If Not (bandhaydatos) Then
                For i = 1 To 10
                    gComisionesJefeB(i).desde = 0
                    gComisionesJefeB(i).hasta = 0
                    gComisionesJefeB(i).Comision = 0
                    gComisionesJefeB(i).ComisionC = 0
                    gComisionesJefeB(i).ComisionSC = 0
                Next i
            End If
        End If
    Case "C"
        If Len(cad) > 0 Then
            bandhaydatos = False
            W = Split(cad, ";")
            For i = 1 To UBound(W)
                    v = Split(W(i - 1), ",")
                    gComisionesJefeC(i).desde = CCur("0" & v(0))
                    gComisionesJefeC(i).hasta = CCur("0" & v(1))
                    gComisionesJefeC(i).Comision = CCur("0" & v(2))
                    gComisionesJefeC(i).ComisionC = CCur("0" & v(3))
                    gComisionesJefeC(i).ComisionSC = CCur("0" & v(4))
                    bandhaydatos = True
            Next i
        Else
            If Not (bandhaydatos) Then
                For i = 1 To 10
                    gComisionesJefeC(i).desde = 0
                    gComisionesJefeC(i).hasta = 0
                    gComisionesJefeC(i).Comision = 0
                    gComisionesJefeC(i).ComisionC = 0
                    gComisionesJefeC(i).ComisionSC = 0
                Next i
            End If
        End If
    Case "D"
        If Len(cad) > 0 Then
            bandhaydatos = False
            W = Split(cad, ";")
            For i = 1 To UBound(W)
                    v = Split(W(i - 1), ",")
                    gComisionesJefeD(i).desde = CCur("0" & v(0))
                    gComisionesJefeD(i).hasta = CCur("0" & v(1))
                    gComisionesJefeD(i).Comision = CCur("0" & v(2))
                    gComisionesJefeD(i).ComisionC = CCur("0" & v(3))
                    gComisionesJefeD(i).ComisionSC = CCur("0" & v(4))
                    bandhaydatos = True
            Next i
        Else
            If Not (bandhaydatos) Then
                For i = 1 To 10
                    gComisionesJefeD(i).desde = 0
                    gComisionesJefeD(i).hasta = 0
                    gComisionesJefeD(i).Comision = 0
                    gComisionesJefeD(i).ComisionC = 0
                    gComisionesJefe(i).ComisionSC = 0
                Next i
            End If
        End If
    End Select
    Exit Sub
ErrTrap:
    'si no existe el archivo coloca ceros
    Exit Sub
End Sub

Public Sub CargaTipoTransxTipoTransTipoComp(ByRef Modulo As String, ByRef lst As ListBox, tipoTrans As String, TipoComp As String)
    Dim rs As Recordset, Vector As Variant
    Dim numMod As Integer, i As Integer
    'Prepara la lista de tipos de transaccion
    lst.Clear
    Vector = Split(Modulo, ",")
    numMod = UBound(Vector, 1)
    If numMod = -1 Then
        Set rs = gobjMain.EmpresaActual.ListaGNTransTipoTRansTipoComp("", False, True, tipoTrans, TipoComp)
        With rs
            If Not (.EOF) Then
                .MoveFirst
                Do Until .EOF
                    lst.AddItem !CodTrans & "  " & !NombreTrans
                    lst.ItemData(lst.NewIndex) = Len(!CodTrans)
                    .MoveNext
                Loop
            End If
        End With
        rs.Close
    Else
        For i = 0 To numMod
            Set rs = gobjMain.EmpresaActual.ListaGNTransTipoTRansTipoComp(CStr(Vector(i)), False, True, tipoTrans, TipoComp)
            With rs
                If Not (.EOF) Then
                    .MoveFirst
                    Do Until .EOF
                        lst.AddItem !CodTrans & "  " & !NombreTrans
                        lst.ItemData(lst.NewIndex) = Len(!CodTrans)
                        .MoveNext
                    Loop
                End If
            End With
            rs.Close
        Next i
    End If
    Set rs = Nothing
End Sub

Public Sub EscribirIntervalosGnOpcionMontoVentas()
    Dim file As String, NumFile As Long, i As Integer
    Dim LINEA As String
    On Error GoTo ErrTrap
    LINEA = ""
    For i = 1 To 10
            LINEA = LINEA & gMonto(i).desde & "," & gMonto(i).hasta & "," & gMonto(i).grupo & ";"
    Next i
     gobjMain.EmpresaActual.GNOpcion.AsignarValor "TablaPCGxMontoVentas", LINEA
     gobjMain.EmpresaActual.GNOpcion.GrabarGNOpcion2
    
    LeerIntervalosGnOpcionMontoVentas
    Exit Sub
    
ErrTrap:
    MsgBox Err.Description
    Exit Sub
End Sub


Public Sub LeerIntervalosGnOpcionMontoVentas()
    Dim cad As String, NumFile As Long, i As Integer, W As Variant
    Dim LINEA As String, v As Variant, bandhaydatos As Boolean
    On Error GoTo ErrTrap
    
    
    cad = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TablaPCGxMontoVentas")
        If Len(cad) > 0 Then
            bandhaydatos = False
            W = Split(cad, ";")
            For i = 1 To UBound(W)
                    v = Split(W(i - 1), ",")
                    gMonto(i).desde = CCur("0" & v(0))
                    gMonto(i).hasta = CCur("0" & v(1))
                    gMonto(i).grupo = v(2)
                    bandhaydatos = True
            Next i
        Else
            If Not (bandhaydatos) Then
                For i = 1 To 10
                    gMonto(i).desde = 0
                    gMonto(i).hasta = 0
                    gMonto(i).grupo = ""
                Next i
            End If
        End If
    Exit Sub
    
ErrTrap:
    'si no existe el archivo coloca ceros

    Exit Sub
End Sub

Public Sub LeerIntervalosGnOpcionMontoVentasCobros()
    Dim cad As String, NumFile As Long, i As Integer, W As Variant
    Dim LINEA As String, v As Variant, bandhaydatos As Boolean
    On Error GoTo ErrTrap
    cad = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TablaPCGxMontoVentasCobro")
        If Len(cad) > 0 Then
            bandhaydatos = False
            W = Split(cad, ";")
            For i = 1 To UBound(W)
                    v = Split(W(i - 1), ",")
                    gMontoCobro(i).desde = CCur("0" & v(0))
                    gMontoCobro(i).hasta = CCur("0" & v(1))
                    gMontoCobro(i).diasMorosidad = CCur("0" & v(2))
                    gMontoCobro(i).grupo = v(3)
                    bandhaydatos = True
            Next i
        Else
            If Not (bandhaydatos) Then
                For i = 1 To 10
                    gMontoCobro(i).desde = 0
                    gMontoCobro(i).hasta = 0
                    gMontoCobro(i).diasMorosidad = 0
                    gMontoCobro(i).grupo = ""
                Next i
            End If
        End If
        Exit Sub
ErrTrap:
    'si no existe el archivo coloca ceros
    Exit Sub
End Sub

Public Sub EscribirIntervalosGnOpcionMontoVentasCobro()
    Dim file As String, NumFile As Long, i As Integer
    Dim LINEA As String
    On Error GoTo ErrTrap
    LINEA = ""
    For i = 1 To 10
            LINEA = LINEA & gMontoCobro(i).desde & "," & gMontoCobro(i).hasta & "," & gMontoCobro(i).diasMorosidad & "," & gMontoCobro(i).grupo & ";"
    Next i
     gobjMain.EmpresaActual.GNOpcion.AsignarValor "TablaPCGxMontoVentasCobro", LINEA
     gobjMain.EmpresaActual.GNOpcion.GrabarGNOpcion2
    LeerIntervalosGnOpcionMontoVentasCobros
    Exit Sub
ErrTrap:
    MsgBox Err.Description
    Exit Sub
End Sub

