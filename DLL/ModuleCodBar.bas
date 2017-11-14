Attribute VB_Name = "ModuleCodBar"
Option Explicit


Public Type T_CONFIG
    CodTransListado As String       'Código de trans. para listado de trans
    CodAltVisible As Boolean        'True visualiza e imprime en los listados y etiquetas
    LetrasClaves As String          'Guardará las letras que van a reemplazar el precio de costo.
    MargenPredeterminado As String  'Guardará toda la configuración de los márgenes y adicionales de la hoja
    ImprimeLineaH As Boolean        'True imprime una línea horizontal entre etiquetas
    ImprimeCosto As Boolean         'True imprime el costo en las etiquetas
    
    '*** MAKOTO 05/dic/00 Agregado
    costeo As Integer               'COSTO_PROM, COSTO_FIFO, COSTO_LIFO, COSTO_ULTIMO
    
    '*** MAKOTO 09/feb/01 Agregado
    PantallaInicial As String       '"PTRANS" / "PITEM"
    
    '*** ANGEL  21/May/01
'    VisualizaPrecios As T_PRECIOS  'Para especificar que precios necesita ver

    '*** ANGEL  10/Ene/02 Agregado
    PrecioEtiqueta As Integer       'Define el precio que se va imprimir en la etiqueta
    
    '*** ANGEL 19/Feb/02 Agregado
    TipoEtiqueta As Boolean         'True = Precios ; False = Código Barras
    
    '*** ANGEL 21/Jul/2003
    PosCampoEMP As String           'Define posiciones del campo empresa en la etiqueta
    PosCampoCOD As String           'Define posiciones del campo codigo en la etiqueta
    PosCampoCA1 As String           'Define posiciones del campo codalterno en la etiqueta
    PosCampoDES As String           'Define posiciones del campo descripcion en la etiqueta
    PosCampoCOS As String           'Define posiciones del campo costo en la etiqueta
    PosCampoPRE1 As String           'Define posiciones del campo precio 1 en la etiqueta
    PosCampoPRE2 As String           'Define posiciones del campo precio 2 en la etiqueta
    PosCampoPRE3 As String           'Define posiciones del campo precio 3 en la etiqueta
    PosCampoPRE4 As String           'Define posiciones del campo precio 4 en la etiqueta
    PosCampoPRE5 As String           'Define posiciones del campo precio 5 en la etiqueta AUC
    PosCampoPIVA1 As String           'Define posiciones del campo precio+ IVA 1 en la etiqueta
    PosCampoPIVA2 As String           'Define posiciones del campo precio+ IVA 2 en la etiqueta
    PosCampoPIVA3 As String           'Define posiciones del campo precio+ IVA 3 en la etiqueta
    PosCampoPIVA4 As String           'Define posiciones del campo precio+ IVA 4 en la etiqueta
    PosCampoPIVA5 As String           'Define posiciones del campo precio+ IVA 5 en la etiqueta AUC
    PosCampoCBR As String           'Define posiciones del campo código de barra en la etiqueta
    
    PosCampoIVG1 As String
    PosCampoIVGD1 As String
    PosCampoIVG2 As String
    PosCampoIVGD2 As String
    PosCampoIVG3 As String
    PosCampoIVGD3 As String
    PosCampoIVG4 As String
    PosCampoIVGD4 As String
    PosCampoIVG5 As String
    PosCampoIVGD5 As String
    PosCampoCant As String
    PosCampoCodigoDer As String
    
    MargenesCodBar As String        'Define posiciones del campo de cod. barras
    ImprimeLineas As Boolean
    UtilizaCostoUC As Boolean
    
    PosCampoAFEMP As String           'Define posiciones del campo empresa en la etiqueta
    PosCampoAFCOD As String           'Define posiciones del campo codigo en la etiqueta
    PosCampoAFCA1 As String           'Define posiciones del campo codalterno en la etiqueta
    PosCampoAFDES As String           'Define posiciones del campo descripcion en la etiqueta
    PosCampoAFCBR As String           'Define posiciones del campo código de barra en la etiqueta
    PosCampoAFDEP As String
    PosCampoAFFCP As String
    MargenesAFCodBar As String        'Define posiciones del campo de cod. barras
    MargenAFPredeterminado As String
End Type

Public Type T_XY
    X As Single
    Y As Single
End Type

Public Type T_DATOS_CODBAR
    NombreEmpresa As T_XY
    codigoItem As T_XY
    CodAlterno As T_XY
    Descripcion As T_XY
    costo As T_XY
    precio1 As T_XY
    precio2 As T_XY
    precio3 As T_XY
    precio4 As T_XY
    precio5 As T_XY
    PrecioIVA1 As T_XY
    PrecioIVA2 As T_XY
    PrecioIVA3 As T_XY
    PrecioIVA4 As T_XY
    PrecioIVA5 As T_XY
    CodigoBarras As T_XY
    IVGrupo1 As T_XY
    IVGrupo1D As T_XY
    IVGrupo2 As T_XY
    IVGrupo2D As T_XY
    IVGrupo3 As T_XY
    IVGrupo3D As T_XY
    IVGrupo4 As T_XY
    IVGrupo4D As T_XY
    IVGrupo5 As T_XY
    IVGrupo5D As T_XY
    Cantidad As T_XY
    CodigoDer As T_XY
End Type
Public gconfig As T_CONFIG
Public gCodBar() As T_DATOS_CODBAR
Public gColumnas As Configuracion   'Para configuración de columnas '*** ANGEL 10/ene/02 Agregado
Public gFmtCosto As String
Public UltCodItemNum As Long





Public Function ObtieneCaminoRelativo(dest As String, orig As String)
'dest : Camino destino. Tiene que terminar con "\".
'orig : Camino de origen. Tiene que terminar con "\".
'
    Dim i As Integer, j As Integer, s As String, n As Integer
    
    For i = 1 To Len(orig)
        If Mid$(orig, i, 1) <> Mid$(dest, i, 1) Or i > Len(dest) Then
            For j = i To Len(orig)
                If Mid$(orig, j, 1) = "\" Then s = s & "..\"
            Next j
            Exit For
        Else
            If Mid$(orig, i, 1) = "\" Then n = i
        End If
    Next i
    
    If n < Len(dest) Then
        If n = Len(orig) Then
            s = "\" & Right$(dest, Len(dest) - n)
        Else
             s = s & Right$(dest, Len(dest) - n)
        End If
    End If
    
    ObtieneCaminoRelativo = s
End Function










Public Sub GenerarBarcode(ByVal bc_string As String, obj As VSPrinter, ByVal X As Single, Y As Single, Alto As Single)

    Dim xpos!, xpos2!, y1!, y2!, dw!, bw!
    Dim new_string$, px!, px2!, n!, c!, bc_pattern$, i!
    Dim ancho As Double

    Dim vp As VSPrinter


    Set vp = obj
    'define barcode patterns
    Dim bc(90) As String
    bc(1) = "1 1221"            'pre-amble
    bc(2) = "1 1221"            'post-amble
    bc(48) = "11 221"           'digits
    bc(49) = "21 112"
    bc(50) = "12 112"
    bc(51) = "22 111"
    bc(52) = "11 212"
    bc(53) = "21 211"
    bc(54) = "12 211"
    bc(55) = "11 122"
    bc(56) = "21 121"
    bc(57) = "12 121"
                                'capital letters
    bc(65) = "211 12"           'A
    bc(66) = "121 12"           'B
    bc(67) = "221 11"           'C
    bc(68) = "112 12"           'D
    bc(69) = "212 11"           'E
    bc(70) = "122 11"           'F
    bc(71) = "111 22"           'G
    bc(72) = "211 21"           'H
    bc(73) = "121 21"           'I
    bc(74) = "112 21"           'J
    bc(75) = "2111 2"           'K
    bc(76) = "1211 2"           'L
    bc(77) = "2211 1"           'M
    bc(78) = "1121 2"           'N
    bc(79) = "2121 1"           'O
    bc(80) = "1221 1"           'P
    bc(81) = "1112 2"           'Q
    bc(82) = "2112 1"           'R
    bc(83) = "1212 1"           'S
    bc(84) = "1122 1"           'T
    bc(85) = "2 1112"           'U
    bc(86) = "1 2112"           'V
    bc(87) = "2 2111"           'W
    bc(88) = "1 1212"           'X
    bc(89) = "2 1211"           'Y
    bc(90) = "1 2211"           'Z
                                'Misc
    bc(32) = "1 2121"           'space
    bc(35) = ""                 '# cannot do!
    bc(36) = "1 1 1 11"         '$
    bc(37) = "11 1 1 1"         '%
    bc(43) = "1 11 1 1"         '+
    bc(45) = "1 1122"           '-
    bc(47) = "1 1 11 1"         '/
    bc(46) = "2 1121"           '.
    bc(64) = ""                 '@ cannot do!
    bc(65) = "1 1221"           '*

    bc_string = UCase(bc_string)

    'dimensions
    With vp
        dw = 1                                      'space between bars
        If dw < 1 Then dw = 1
        new_string = Chr$(1) & bc_string & Chr$(2)  'add pre-amble, post-amble

        y1 = Y                           'Punto y donde empieza la barra
        y2 = Y + Alto                    'Define el alto de la barra

        'Recupera el ancho de barra, dependiendo de la intensidad de las barras se realiza la lectura
        bw = GetSetting(APPNAME, "SiiPrecioA", "AnchoBarra", 2)
        If bw < 2 Then bw = 2
        'jeaa 24/05/2005
        ancho = GetSetting(APPNAME, "SiiPrecioA", "AnchoTotalCB", 1.2)
        px = (.TwipsPerPixelX + bw) / ancho '2.2 De estos valores depende el ancho de barras, antes 3
        px2 = (px * (bw - 1)) / 2           '1.2 De estos valores depende el ancho de barras, antes 2
        px2 = 2
        'draw each character in barcode string
        xpos = X
        dw = Round(dw * px, 1)

        For n = 1 To Len(new_string)
            c = Asc(Mid$(new_string, n, 1))
            If c > 90 Then c = 0
            bc_pattern$ = bc(c)

            'draw each bar
            For i = 1 To Len(bc_pattern$)
                Select Case Mid$(bc_pattern$, i, 1)
                Case " "
                    'space
                    vp.PenColor = &HFFFFFF
                    vp.BrushColor = &HFFFFFF
                    xpos2 = xpos + Round(px * dw, 1)
                    vp.DrawRectangle xpos, y1, xpos2, y2
                    xpos = xpos2 'xpos + dw + 10

                Case "1"
                    'space
                    vp.PenColor = &HFFFFFF
                    vp.BrushColor = &HFFFFFF
                    xpos2 = xpos + Round(px * dw, 1)
                    vp.DrawRectangle xpos, y1, xpos2, y2
                    xpos = xpos2 'xpos + dw + 10
                    'line
                    vp.PenColor = &H0
                    vp.BrushColor = &H0
                    xpos2 = xpos + Round(px * dw, 1)
                    vp.DrawRectangle xpos, y1, xpos2, y2
                    xpos = xpos2 'xpos + dw + 10

                Case "2"
                    'space
                    vp.PenColor = &HFFFFFF
                    vp.BrushColor = &HFFFFFF
                    xpos2 = xpos + Round(px * dw, 1)
                    vp.DrawRectangle xpos, y1, xpos2, y2
                    xpos = xpos2 'xpos + dw + 10
                    'wide line
                    vp.PenColor = &H0
                    vp.BrushColor = &H0
                    xpos2 = xpos + px2 * Round((dw * px), 1)
                    vp.DrawRectangle xpos, y1, xpos2, y2
                    xpos = xpos2 'xpos + (px2 * dw * px) + 10
                End Select
            Next
        Next


        '*** Revisar bien si hace algún efecto o no 26/02/2002 Angel
        '1 more space
        vp.PenColor = &HFFFFFF
        vp.BrushColor = &HFFFFFF
        xpos2 = xpos + Round(px * dw, 1)
        vp.DrawRectangle xpos, y1, xpos2, y2
        xpos = xpos2
    End With
End Sub

Public Sub RecuperaConfig()
    Dim valorpredeterminado As String, clavepredeterminada As String
    On Error GoTo ErrTrap
    valorpredeterminado = "5,10,13,13,13,13,210,297,38,25,0,0,0,0"
    clavepredeterminada = "A,B,C,D,E,F,G,H,I,J"
    With gconfig
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_CodTransListado")) > 0 Then
            .CodTransListado = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_CodTransListado")
        Else
            .CodTransListado = "CP"
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_CodAltVisible")) > 0 Then
            .CodAltVisible = IIf(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_CodAltVisible") = "1", True, False)
        Else
            .CodAltVisible = False
        End If
        
                
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_LetrasClaves")) > 0 Then
            .LetrasClaves = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_LetrasClaves")
        Else
            .LetrasClaves = clavepredeterminada
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_MargenPredeterminado")) > 0 Then
            .MargenPredeterminado = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_MargenPredeterminado")
        Else
            .MargenPredeterminado = valorpredeterminado
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_ImprimeLineaH")) > 0 Then
            .ImprimeLineaH = IIf(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_ImprimeLineaH") = "1", True, False)
        Else
            .ImprimeLineaH = True
        End If

        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_ImprimeCosto")) > 0 Then
            .ImprimeCosto = IIf(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_ImprimeCosto") = "1", True, False)
        Else
            .ImprimeCosto = True
        End If

        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_costeo")) > 0 Then
            .costeo = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_costeo")
        Else
            .costeo = COSTO_PROM
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PantallaInicial")) > 0 Then
            .PantallaInicial = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PantallaInicial")
        Else
            .PantallaInicial = "PTRANS"
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PrecioEtiqueta")) > 0 Then
            .PrecioEtiqueta = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PrecioEtiqueta")
        Else
            .PrecioEtiqueta = 1
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_TipoEtiqueta")) > 0 Then
            .TipoEtiqueta = IIf(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_TipoEtiqueta") = "1", True, False)
        Else
            .TipoEtiqueta = True
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoEMP")) > 0 Then
            .PosCampoEMP = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoEMP")
        Else
            .PosCampoEMP = "2;2;Arial;12;1;15;;"
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoCOD")) > 0 Then
            .PosCampoCOD = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoCOD")
        Else
            .PosCampoCOD = "2;6;Arial;10;1;5;;"
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoCA1")) > 0 Then
            .PosCampoCA1 = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoCA1")
        Else
            .PosCampoCA1 = "20;6;Arial;10;0;10;;"
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoDES")) > 0 Then
            .PosCampoDES = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoDES")
        Else
            .PosCampoDES = "2;10;Arial;10;1;20;;"
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoCOS")) > 0 Then
            .PosCampoCOS = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoCOS")
        Else
            .PosCampoCOS = "2;10;Arial;10;0;0;#,0.00;"
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoPRE1")) > 0 Then
            .PosCampoPRE1 = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoPRE1")
        Else
            .PosCampoPRE1 = "2;8;Arial;12;0;0;#,0.00;"
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoPRE2")) > 0 Then
            .PosCampoPRE2 = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoPRE2")
        Else
            .PosCampoPRE2 = "2;8;Arial;12;0;0;#,0.00;"
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoPRE3")) > 0 Then
            .PosCampoPRE3 = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoPRE3")
        Else
            .PosCampoPRE3 = "2;8;Arial;12;0;0;#,0.00;"
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoPRE4")) > 0 Then
            .PosCampoPRE4 = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoPRE4")
        Else
            .PosCampoPRE4 = "2;8;Arial;12;0;0;#,0.00;"
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoPRE5")) > 0 Then
            .PosCampoPRE5 = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoPRE5")
        Else
            .PosCampoPRE5 = "2;8;Arial;12;0;0;#,0.00;"
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoPIVA1")) > 0 Then
            .PosCampoPIVA1 = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoPIVA1")
        Else
            .PosCampoPIVA1 = "2;8;Arial;12;0;0;#,0.00;"
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoPIVA2")) > 0 Then
            .PosCampoPIVA2 = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoPIVA2")
        Else
            .PosCampoPIVA2 = "2;8;Arial;12;0;0;#,0.00;"
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoPIVA3")) > 0 Then
            .PosCampoPIVA3 = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoPIVA3")
        Else
            .PosCampoPIVA3 = "2;8;Arial;12;0;0;#,0.00;"
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoPIVA4")) > 0 Then
            .PosCampoPIVA4 = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoPIVA4")
        Else
            .PosCampoPIVA4 = "2;8;Arial;12;0;0;#,0.00;"
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoPIVA5")) > 0 Then
            .PosCampoPIVA5 = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoPIVA5")
        Else
            .PosCampoPIVA5 = "2;8;Arial;12;0;0;#,0.00;"
        End If

        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoCBR")) > 0 Then
            .PosCampoCBR = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoCBR")
        Else
            .PosCampoCBR = "2;10;C39HrP24DhTt;8;0;0;;"
        End If

        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_MargenesCodBar")) > 0 Then
            .MargenesCodBar = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_MargenesCodBar")
        Else
            .MargenesCodBar = "5,10,290,210,20,20,14,14,8,0.02,0,0"
        End If

        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_ImprimeLineas")) > 0 Then
            .ImprimeLineas = IIf(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_ImprimeLineas") = "1", True, False)
        Else
            .ImprimeLineas = True
        End If
        
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_UtilizaCostoUC")) > 0 Then
            .UtilizaCostoUC = IIf(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_UtilizaCostoUC") = "1", True, False)
        Else
            .UtilizaCostoUC = False
        End If

'AUC 04/2012
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoIVG1")) > 0 Then
            .PosCampoIVG1 = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoIVG1")
        Else
            .PosCampoIVG1 = "2;8;Arial;12;0;0;#,0.00;"
        End If
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoIVGD1")) > 0 Then
            .PosCampoIVGD1 = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoIVGD1")
        Else
            .PosCampoIVGD1 = "2;8;Arial;12;0;0;#,0.00;"
        End If
        
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoIVG2")) > 0 Then
            .PosCampoIVG2 = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoIVG2")
        Else
            .PosCampoIVG2 = "2;8;Arial;12;0;0;#,0.00;"
        End If
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoIVGD2")) > 0 Then
            .PosCampoIVGD2 = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoIVGD2")
        Else
            .PosCampoIVGD2 = "2;8;Arial;12;0;0;#,0.00;"
        End If
        

        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoIVG3")) > 0 Then
            .PosCampoIVG3 = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoIVG3")
        Else
            .PosCampoIVG3 = "2;8;Arial;12;0;0;#,0.00;"
        End If
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoIVGD3")) > 0 Then
            .PosCampoIVGD3 = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoIVGD3")
        Else
            .PosCampoIVGD3 = "2;8;Arial;12;0;0;#,0.00;"
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoIVG4")) > 0 Then
            .PosCampoIVG4 = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoIVG4")
        Else
            .PosCampoIVG4 = "2;8;Arial;12;0;0;#,0.00;"
        End If
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoIVGD4")) > 0 Then
            .PosCampoIVGD4 = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoIVGD4")
        Else
            .PosCampoIVGD4 = "2;8;Arial;12;0;0;#,0.00;"
        End If
        
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoIVG5")) > 0 Then
            .PosCampoIVG5 = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoIVG5")
        Else
            .PosCampoIVG5 = "2;8;Arial;12;0;0;#,0.00;"
        End If
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoIVGD5")) > 0 Then
            .PosCampoIVGD5 = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoIVGD5")
        Else
            .PosCampoIVGD5 = "2;8;Arial;12;0;0;#,0.00;"
        End If
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoCant")) > 0 Then
            .PosCampoCant = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoCant")
        Else
            .PosCampoCant = "2;8;Arial;12;0;0;#,0.00;"
        End If
         If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoCodigoDer")) > 0 Then
            .PosCampoCodigoDer = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoCodigoDer")
        Else
            .PosCampoCodigoDer = "2;8;Arial;12;0;0;#,0.00;"
        End If

''------------------------- acTIVOS fIJOS

        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoAFEMP")) > 0 Then
            .PosCampoAFEMP = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoAFEMP")
        Else
            .PosCampoAFEMP = "2;2;Arial;12;1;15;;"
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoAFCOD")) > 0 Then
            .PosCampoAFCOD = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoAFCOD")
        Else
            .PosCampoAFCOD = "2;6;Arial;10;1;5;;"
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoAFCA1")) > 0 Then
            .PosCampoAFCA1 = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoAFCA1")
        Else
            .PosCampoAFCA1 = "20;6;Arial;10;0;10;;"
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoAFDES")) > 0 Then
            .PosCampoAFDES = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoAFDES")
        Else
            .PosCampoAFDES = "2;10;Arial;10;1;20;;"
        End If
        

        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoAFCBR")) > 0 Then
            .PosCampoAFCBR = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_PosCampoAFCBR")
        Else
            .PosCampoAFCBR = "2;10;C39HrP24DhTt;8;0;0;;"
        End If

        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_MargenesAFCodBar")) > 0 Then
            .MargenesAFCodBar = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_MargenesAFCodBar")
        Else
            .MargenesAFCodBar = "5,10,290,210,20,20,14,14,8,0.02,0,0"
        End If


        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_MargenesAFDEP")) > 0 Then
            .PosCampoAFDEP = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_MargenesAFDEP")
        Else
            .PosCampoAFDEP = "2;10;Arial;10;1;20;;"
        End If


        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_MargenesAFFCP")) > 0 Then
            .PosCampoAFFCP = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_MargenesAFDFCP")
        Else
            .PosCampoAFFCP = "2;10;Arial;10;1;20;;"
        End If


        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_MargenAFPredeterminado")) > 0 Then
            .MargenPredeterminado = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_MargenAFPredeterminado")
        Else
            .MargenAFPredeterminado = valorpredeterminado
        End If
    End With
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub


