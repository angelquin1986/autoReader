Attribute VB_Name = "Impresion"
Option Explicit

Type DATOS_VxTxCABINA
    cabina As String
    Trafico As String
    Destino As String
    TotMinuto As Currency
    Neto As Currency
    ICE As Currency
    IVA As Currency
    Total As Currency
End Type

Public Matriz() As DATOS_VxTxCABINA

'***Agregado. julio/2003. Angel. Para poder imprimir reporte específico
'Public Sub GeneralImprimeModGraf(ByVal grd As Control, _
'                                                                titulo As String, desde As Date, hasta As Date)
'
'    Dim vp As SiiPrint.PreviewVSPrint
'    Dim Cadena As String, v_datos As Variant, titul As String
'    Dim fmt_can As String, fmt_ctn As String, fmt_vtb As String, fmt_vol As String
'
'    Cadena = desde & " - " & hasta
''    v_datos = GrillaProcesada(grd)
'
'    Set vp = New SiiPrint.PreviewVSPrint
'    vp.Caption = "Impresión de Reporte: " & titulo
'    GenerarDoc vp, v_datos, titulo, Cadena
'    vp.ShowModal
'End Sub

Private Function GrillaProcesada(ByVal grd As VSFlexGrid) As Variant
    Dim cod_zona As String, cod_cli As String, cod_cat As String, v() As Variant
    Dim num_cabina As String, Trafico As String, Destino As String, NumMin As String
    Dim i As Long, j As Long, l As Long, cont As Long, band As Boolean
    Dim fmt As String, fmt_c As String
    
    Const COL_NUM = 0
    Const COL_CABINA = 1
    Const COL_TRAFICO = 2
    Const COL_DESTINO = 3
    Const COL_TOTMINUTOS = 4
    Const COL_NETO = 5
    Const COL_ICE = 6
    Const COL_IVA = 7
    Const COL_TOTAL = 8
   
    
    cont = 0
    ReDim Preserve Matriz(cont)
    num_cabina = Trim$(grd.TextMatrix(grd.FixedRows, COL_CABINA))
    Matriz(cont).cabina = Trim$(grd.TextMatrix(grd.FixedRows, COL_CABINA))
    
    cont = cont + 1
    ReDim Preserve Matriz(cont)
    Trafico = Trim$(grd.TextMatrix(grd.FixedRows, COL_TRAFICO))
    Matriz(cont).Trafico = Trim$(grd.TextMatrix(grd.FixedRows, COL_TRAFICO))
    
    Destino = Trim$(grd.TextMatrix(grd.FixedRows, COL_DESTINO))
    With grd
        For i = .FixedRows To .Rows - 1
            band = False
            If .IsSubtotal(i) = True Then
                cont = cont + 1
                ReDim Preserve Matriz(cont)
                     If Len(Matriz(cont - 1).Trafico) = 0 Then
                            Matriz(cont).Trafico = Trim$(grd.TextMatrix(i, COL_TRAFICO))
                            If Len(Matriz(cont).Trafico) = 0 And Len(Matriz(cont - 1).cabina) = 0 Then
                                Matriz(cont).cabina = Trim$(grd.TextMatrix(i, COL_CABINA))
                            End If
                    End If
                If Len(Destino) > 0 Then
                    Matriz(cont).Destino = Trim$(grd.TextMatrix(i - 1, COL_DESTINO))
                Else
                    ReDim Preserve Matriz(cont)
                    Select Case i
                    Case .Rows - 1
                        Matriz(cont).Total = "TOTAL GENERAL: "
                    Case Else
                        MsgBox Matriz(cont).Trafico
            
                    End Select
                End If
                
                fmt_c = "#,0.0000"
                Matriz(cont).Destino = .TextMatrix(i, COL_DESTINO)
                Matriz(cont).TotMinuto = .TextMatrix(i, COL_TOTMINUTOS)
                Matriz(cont).Neto = .TextMatrix(i, COL_NETO)
                Matriz(cont).ICE = .TextMatrix(i, COL_ICE)
                Matriz(cont).IVA = .TextMatrix(i, COL_IVA)
                Matriz(cont).Total = .TextMatrix(i, COL_TOTAL)
                band = True
                If i < .Rows - 1 Then Destino = Trim$(grd.TextMatrix(grd.FixedRows, COL_DESTINO))
            Else
                If Trim$(.TextMatrix(i, COL_CABINA)) <> num_cabina Then
               
                    ReDim Preserve Matriz(cont)
                    num_cabina = Trim$(grd.TextMatrix(i, COL_CABINA))
                    Matriz(cont).cabina = Trim$(grd.TextMatrix(i, COL_CABINA))
                    cont = cont + 1
                    ReDim Preserve Matriz(cont)
                    Trafico = Trim$(grd.TextMatrix(i, COL_TRAFICO))
                    Matriz(cont).Trafico = Trim$(grd.TextMatrix(i, COL_TRAFICO))
                    cont = cont + 1
                    ReDim Preserve Matriz(cont)
                        Matriz(cont).Destino = .TextMatrix(i, COL_DESTINO)
                         Matriz(cont).TotMinuto = .TextMatrix(i, COL_TOTMINUTOS)
                         Matriz(cont).Neto = .TextMatrix(i, COL_NETO)
                         Matriz(cont).ICE = .TextMatrix(i, COL_ICE)
                         Matriz(cont).IVA = .TextMatrix(i, COL_IVA)
                         Matriz(cont).Total = .TextMatrix(i, COL_TOTAL)
                Else
                    If Trim$(.TextMatrix(i, COL_TRAFICO)) <> Trafico Then
                        cont = cont + 1
                        ReDim Preserve Matriz(cont)
                       Trafico = Trim$(grd.TextMatrix(i, COL_TRAFICO))
                        Matriz(cont).Trafico = Trim$(grd.TextMatrix(i, COL_TRAFICO))
                        cont = cont + 1
                        ReDim Preserve Matriz(cont)
                        Matriz(cont).Destino = .TextMatrix(i, COL_DESTINO)
                         Matriz(cont).TotMinuto = .TextMatrix(i, COL_TOTMINUTOS)
                         Matriz(cont).Neto = .TextMatrix(i, COL_NETO)
                         Matriz(cont).ICE = .TextMatrix(i, COL_ICE)
                         Matriz(cont).IVA = .TextMatrix(i, COL_IVA)
                         Matriz(cont).Total = .TextMatrix(i, COL_TOTAL)
                    Else
                         cont = cont + 1
                         ReDim Preserve Matriz(cont)
                         Matriz(cont).Destino = .TextMatrix(i, COL_DESTINO)
                         Matriz(cont).TotMinuto = .TextMatrix(i, COL_TOTMINUTOS)
                         Matriz(cont).Neto = .TextMatrix(i, COL_NETO)
                         Matriz(cont).ICE = .TextMatrix(i, COL_ICE)
                         Matriz(cont).IVA = .TextMatrix(i, COL_IVA)
                         Matriz(cont).Total = .TextMatrix(i, COL_TOTAL)
                    End If
                End If
            End If
            If band = True Then cont = cont + 1
        Next i
    End With
'
    
    For i = LBound(Matriz, 1) To UBound(Matriz, 1)
        ReDim Preserve v(8, i)
        v(0, i) = Matriz(i).cabina
        v(1, i) = Matriz(i).Trafico
        v(2, i) = Matriz(i).Destino
        If Len(v(2, i)) <> 0 Then
            v(3, i) = Format(IIf(Matriz(i).TotMinuto = 0, "0.0000", Matriz(i).TotMinuto), fmt_c)
            v(4, i) = Format(IIf(Matriz(i).Neto = 0, "0.0000", Matriz(i).Neto), fmt_c)
            v(5, i) = Format(IIf(Matriz(i).ICE = 0, "0.0000", Matriz(i).ICE), fmt_c)
            v(6, i) = Format(IIf(Matriz(i).IVA = 0, "0.0000", Matriz(i).IVA), fmt_c)
            v(7, i) = Format(IIf(Matriz(i).Total = 0, "0.0000", Matriz(i).Total), fmt_c)
            
        Else
            v(3, i) = Format(IIf(Matriz(i).TotMinuto = 0, "", Matriz(i).TotMinuto), fmt_c)
            v(4, i) = Format(IIf(Matriz(i).Neto = 0, "", Matriz(i).Neto), fmt_c)
            v(5, i) = Format(IIf(Matriz(i).ICE = 0, "", Matriz(i).ICE), fmt_c)
            v(6, i) = Format(IIf(Matriz(i).IVA = 0, "", Matriz(i).IVA), fmt_c)
            v(7, i) = Format(IIf(Matriz(i).Total = 0, "", Matriz(i).Total), fmt_c)
        End If
    Next i
    GrillaProcesada = v()
End Function

''Private Sub GenerarDoc(ByVal vp As SiiPrint.PreviewVSPrint, _
''                       ByVal v As Variant, _
''                       ByVal titulo As String, _
''                       ByVal rango_fecha As String)
''    Dim vpReport As VSPrinter, i As Long
''    'Dim celda_zon As String, celda_cli As String, celda_cat As String
''    Dim celda_cab As String, celda_tra As String, celda_des As String
''    Const COL_NUM = 0
''    Const COL_CABINA = 1
''    Const COL_TRAFICO = 2
''    Const COL_DESTINO = 3
''    Const COL_TOTMINUTOS = 4
''    Const COL_NETO = 5
''    Const COL_ICE = 6
''    Const COL_IVA = 7
''    Const COL_TOTAL = 8
''    Set vpReport = vp.VSPrinter
''    With vpReport
''        .ShowGuides = gdShow
''        .Clear
''        .PhysicalPage = True
''        .MarginHeader = 800
''        .MarginTop = 800
''        .MarginBottom = 1200
''        .MarginFooter = 1200
''        .MarginLeft = 1200
''        .MarginRight = 800
''
''        .PaperSize = pprA4
''        .Orientation = orPortrait
''
''        .StartDoc
''        GeneraCabeceraGrafico vpReport, 100, titulo, rango_fecha
''        .FontName = "ARIAL"
''        .FontSize = 8
''        .Footer = "Fecha Impresión: " & Now & " USUARIO: " & _
''                  gobjMain.UsuarioActual.NombreUsuario & "||Pag: " & "%d"
''        .FontSize = 8
''        .FontBold = False
''        .CurrentY = 1200
''            .StartTable
''                .TableBorder = tbBoxColumns
''                .AddTableArray "<800|2000|<2000|>1100|>1100|>1100|>1100|>1100", _
''                          "CABINA|TRAFICO|DESTINO|TOTAL MINUTOS|VALOR NETO|VALOR   ICE |VALOR   IVA |VALOR TOTAL", v
''                .TableCell(tcFontBold, 0, 1, 0, 8) = True
''
''                For i = 0 To UBound(v, 2) + 1
''                    celda_cab = .TableCell(tcText, i, COL_CABINA, i, COL_CABINA)
''                    celda_tra = .TableCell(tcText, i, COL_TRAFICO, i, COL_TRAFICO)
''                    celda_des = .TableCell(tcText, i, COL_DESTINO, i, COL_DESTINO)
''                    If Len(celda_cab) > 0 Then
''                        .TableCell(tcFontBold, i, COL_NUM, i, COL_DESTINO) = True
''                        .TableCell(tcRowBorder, i, COL_NUM, i, COL_TOTAL) = tbAll
''                    End If
''                    If Len(celda_cab) > 0 Then .TableCell(tcFontBold, i, COL_TRAFICO, i, COL_TOTAL) = True
''                    If InStr(1, celda_tra, "Total", vbBinaryCompare) <> 0 Then
''                        .TableCell(tcFontBold, i, COL_NUM, i, COL_TOTAL) = True
''                        .TableCell(tcFontUnderline, i - 1, COL_DESTINO + 1, i - 1, COL_TOTAL) = True
''                    End If
''                Next i
''            .EndTable
''        .EndDoc
''    End With
''End Sub

Private Sub GeneraCabeceraGrafico(ByRef vp As VSPrinter, y As Single, titulo As String, rango_fecha As String)
    'Genera Cabecera  en Modo Grafico
    Dim antx  As Single
    With vp
        .FontName = "ARIAL"
        .FontSize = 16
        .CurrentY = y
        antx = .CurrentX
        .Text = UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa)
        .FontSize = 12
        .CurrentX = antx
        .CurrentY = .CurrentY + 380
        .Text = "Reporte de " & titulo & vbCrLf & _
                "Fecha de Corte: " & Trim$(rango_fecha)
    End With
End Sub

Private Function ResumenGrillaProcesada(ByVal grd As VSFlexGrid) As Variant
    Dim cod_zona As String, cod_cli As String, cod_cat As String, v() As Variant
    Dim num_cabina As String, Trafico As String, Destino As String, NumMin As String
    Dim i As Long, j As Long, l As Long, cont As Long, band As Boolean
    Dim fmt As String, fmt_c As String
    
    Const COL_NUM = 0
    Const COL_CABINA = 1
    Const COL_TRAFICO = 2
    Const COL_DESTINO = 3
    Const COL_TOTMINUTOS = 4
    Const COL_NETO = 5
    Const COL_ICE = 6
    Const COL_IVA = 7
    Const COL_TOTAL = 8
   
    
    cont = 0
    ReDim Preserve Matriz(cont)
    num_cabina = Trim$(grd.TextMatrix(grd.FixedRows, COL_CABINA))
    Matriz(cont).cabina = Trim$(grd.TextMatrix(grd.FixedRows, COL_CABINA))
    
    cont = cont + 1
    ReDim Preserve Matriz(cont)
    
    Destino = Trim$(grd.TextMatrix(grd.FixedRows, COL_DESTINO))
    With grd
        For i = .FixedRows To .Rows - 1
            band = False
            ReDim Preserve Matriz(cont)
            If Trim$(.TextMatrix(i, COL_CABINA)) <> num_cabina Then
                    ReDim Preserve Matriz(cont)
                    num_cabina = Trim$(grd.TextMatrix(i, COL_CABINA))
                    Matriz(cont).cabina = Trim$(grd.TextMatrix(i, COL_CABINA))
                    cont = cont + 1
                    ReDim Preserve Matriz(cont)
            End If
            If .IsSubtotal(i) = True Then
                 If Len(Matriz(cont).Trafico) = 0 Then
                        Matriz(cont).Trafico = Trim$(grd.TextMatrix(i, COL_TRAFICO))
                        If Len(Matriz(cont).Trafico) = 0 And Len(Matriz(cont - 1).cabina) = 0 Then
                            Matriz(cont).cabina = Trim$(grd.TextMatrix(i, COL_CABINA))
                        End If
                End If
                If Len(Destino) > 0 Then
                    Matriz(cont).Destino = Trim$(grd.TextMatrix(i - 1, COL_DESTINO))
                Else
                    ReDim Preserve Matriz(cont)
                    Select Case i
                    Case .Rows - 1
                        Matriz(cont).Total = "TOTAL GENERAL: "
                    Case Else
                        MsgBox Matriz(cont).Trafico
                    End Select
                End If
                fmt_c = "#,0.0000"
                Matriz(cont).Destino = .TextMatrix(i, COL_DESTINO)
                Matriz(cont).TotMinuto = .TextMatrix(i, COL_TOTMINUTOS)
                Matriz(cont).Neto = .TextMatrix(i, COL_NETO)
                Matriz(cont).ICE = .TextMatrix(i, COL_ICE)
                Matriz(cont).IVA = .TextMatrix(i, COL_IVA)
                Matriz(cont).Total = .TextMatrix(i, COL_TOTAL)
                band = True
                If i < .Rows - 1 Then Destino = Trim$(grd.TextMatrix(grd.FixedRows, COL_DESTINO))
            If band = True Then cont = cont + 1
            End If
        Next i
    End With
'
    
    For i = LBound(Matriz, 1) To UBound(Matriz, 1)
        ReDim Preserve v(8, i)
        v(0, i) = Matriz(i).cabina
        If InStr(1, Matriz(i).Trafico, "Total", vbBinaryCompare) <> 0 Then
                v(1, i) = Mid(Matriz(i).Trafico, 6, Len(Matriz(i).Trafico))
        Else
            v(1, i) = Matriz(i).Trafico
        End If
        v(2, i) = Matriz(i).Destino
        
        v(3, i) = Format(IIf(Matriz(i).TotMinuto = 0, "", Matriz(i).TotMinuto), fmt_c)
        If Len(v(3, i)) <> 0 Then
            v(4, i) = Format(IIf(Matriz(i).Neto = 0, "0.0000", Matriz(i).Neto), fmt_c)
            v(5, i) = Format(IIf(Matriz(i).ICE = 0, "0.0000", Matriz(i).ICE), fmt_c)
            v(6, i) = Format(IIf(Matriz(i).IVA = 0, "0.0000", Matriz(i).IVA), fmt_c)
            v(7, i) = Format(IIf(Matriz(i).Total = 0, "0.0000", Matriz(i).Total), fmt_c)
            
        Else
            v(4, i) = Format(IIf(Matriz(i).Neto = 0, "", Matriz(i).Neto), fmt_c)
            v(5, i) = Format(IIf(Matriz(i).ICE = 0, "", Matriz(i).ICE), fmt_c)
            v(6, i) = Format(IIf(Matriz(i).IVA = 0, "", Matriz(i).IVA), fmt_c)
            v(7, i) = Format(IIf(Matriz(i).Total = 0, "", Matriz(i).Total), fmt_c)
        End If
    Next i
    
    ResumenGrillaProcesada = v()
End Function

'***Agregado. julio/2003. Angel. Para poder imprimir reporte específico
''Public Sub ResumenImprimeModGraf(ByVal grd As Control, _
''                                                                titulo As String, desde As Date, hasta As Date)
''
''    Dim vp As SiiPrint.PreviewVSPrint
''    Dim Cadena As String, v_datos As Variant, titul As String
''    Dim fmt_can As String, fmt_ctn As String, fmt_vtb As String, fmt_vol As String
''
''    Cadena = desde & " - " & hasta
'''    v_datos = ResumenGrillaProcesada(grd)
''
''    Set vp = New SiiPrint.PreviewVSPrint
''    vp.Caption = "Impresión de Reporte: " & titulo
''    GenerarDocF101Resumen vp, grd, titulo, Cadena
''    vp.ShowModal
''End Sub


'''Private Sub GenerarDocResumen(ByVal vp As SiiPrint.PreviewVSPrint, _
'''                       ByVal v As Variant, _
'''                       ByVal titulo As String, _
'''                       ByVal rango_fecha As String)
'''    Dim vpReport As VSPrinter, i As Long
'''    'Dim celda_zon As String, celda_cli As String, celda_cat As String
'''    Dim celda_cab As String, celda_tra As String, celda_des As String
'''    Const COL_NUM = 0
'''    Const COL_CABINA = 1
'''    Const COL_TRAFICO = 2
'''    Const COL_DESTINO = 3
'''    Const COL_TOTMINUTOS = 4
'''    Const COL_NETO = 5
'''    Const COL_ICE = 6
'''    Const COL_IVA = 7
'''    Const COL_TOTAL = 8
'''    Set vpReport = vp.VSPrinter
'''    With vpReport
'''        .ShowGuides = gdShow
'''        .Clear
'''        .PhysicalPage = True
'''        .MarginHeader = 800
'''        .MarginTop = 800
'''        .MarginBottom = 1200
'''        .MarginFooter = 1200
'''        .MarginLeft = 1200
'''        .MarginRight = 800
'''
'''        .PaperSize = pprA4
'''        .Orientation = orPortrait
'''
'''        .StartDoc
'''        GeneraCabeceraGrafico vpReport, 100, titulo, rango_fecha
'''        .FontName = "ARIAL"
'''        .FontSize = 8
'''        .Footer = "Fecha Impresión: " & Now & " USUARIO: " & _
'''                  gobjMain.UsuarioActual.NombreUsuario & "||Pag: " & "%d"
'''        .FontSize = 8
'''        .FontBold = False
'''        .CurrentY = 1200
'''            .StartTable
'''                .TableBorder = tbBoxColumns
'''                .AddTableArray "<1300|2400|<0|>1200|>1200|>1200|>1200|>1200", _
'''                          "CABINA|TRAFICO|DESTINO|TOTAL MINUTOS|VALOR NETO|VALOR   ICE |VALOR   IVA |VALOR TOTAL", v
'''                .TableCell(tcFontBold, 0, 1, 0, 8) = True
'''                For i = 0 To UBound(v, 2) + 1
'''                    celda_cab = .TableCell(tcText, i, COL_CABINA, i, COL_CABINA)
'''                    celda_tra = .TableCell(tcText, i, COL_TRAFICO, i, COL_TRAFICO)
'''                    celda_des = .TableCell(tcText, i, COL_DESTINO, i, COL_DESTINO)
'''                    If Len(celda_cab) > 0 Then
'''                        .TableCell(tcFontBold, i, COL_NUM, i, COL_TOTAL) = True
'''                        .TableCell(tcRowBorder, i, COL_NUM, i, COL_TOTAL) = tbAll
'''                    End If
'''                    If Len(celda_cab) > 0 Then .TableCell(tcFontBold, i, COL_TRAFICO, i, COL_TOTAL) = True
'''                    If InStr(1, celda_cab, "Total", vbBinaryCompare) <> 0 Then
'''                        .TableCell(tcFontBold, i + 1, COL_NUM, i + 1, COL_TOTAL) = True
'''                    End If
'''                Next i
'''            .EndTable
'''        .EndDoc
'''    End With
'''End Sub

'''Public Sub GeneralImprimeModGrafF101(ByVal grd As Control, _
'''                                                                titulo As String, desde As Date, hasta As Date)
'''
'''    Dim vp As SiiPrint.PreviewVSPrint
'''    Dim Cadena As String, v_datos As Variant, titul As String
'''    Dim fmt_can As String, fmt_ctn As String, fmt_vtb As String, fmt_vol As String
'''
'''    Cadena = desde & " - " & hasta
''''    v_datos = GrillaProcesada(grd)
'''
'''    Set vp = New SiiPrint.PreviewVSPrint
'''    vp.Caption = "Impresión de Reporte: " & titulo
'''    GenerarDocF101 vp, grd, titulo, Cadena
'''    vp.ShowModal
'''End Sub


''Private Sub GenerarDocF101(ByVal vp As SiiPrint.PreviewVSPrint, _
''                       ByVal grd As Control, _
''                       ByVal titulo As String, _
''                       ByVal rango_fecha As String)
''    Dim vpReport As VSPrinter, i As Long
''    'Dim celda_zon As String, celda_cli As String, celda_cat As String
''    Dim celda_cab As String, celda_tra As String, celda_des As String
''          Dim hd, j%, k%, fmt, bd As String
''    Const COL_NUM = 0
''    Const COL_TIPO = 1
''    Const COL_CAMPO = 2
''    Const COL_TOTAL = 3
''    Set vpReport = vp.VSPrinter
''    With vpReport
''        .ShowGuides = gdShow
''        .Clear
''        .PhysicalPage = True
''        .MarginHeader = 800
''        .MarginTop = 800
''        .MarginBottom = 1200
''        .MarginFooter = 1200
''        .MarginLeft = 1200
''        .MarginRight = 800
''
''        .PaperSize = pprA4
''        .Orientation = orPortrait
''
''        .StartDoc
''        GeneraCabeceraGrafico vpReport, 100, titulo, rango_fecha
''        .FontName = "ARIAL"
''        .FontSize = 8
''        .Footer = "Fecha Impresión: " & Now & " USUARIO: " & _
''                  gobjMain.UsuarioActual.NombreUsuario & "||Pag: " & "%d"
''        .FontSize = 8
''        .FontBold = False
''        .CurrentY = 1200
''            .StartTable
''            .TableBorder = tbAll
''            .Paragraph = ""
''            i = 1
''
''                fmt = ">600|<1200|<1200|<1500|<3800|>1500"
''                hd = "#|TIPO|CAMPO|CUENTA|NOMBRE|TOTAL"
''            Do While i <> grd.Rows
''                  grd.Select i, 1
'''                  If GRD.IsSubtotal(i) Then
''                  For j = 0 To 5
''                        grd.Select i, j
''                        Select Case j
''                              Case 4, 3, 2, 1, 0:
''                                          bd = bd & grd.Text & "|"
''                              Case 5:
''                                    If grd.Text <> 0 Then
''                                          bd = bd & Format(grd.Text, "#,###0.00") & "|"
''                                    Else
''                                          bd = bd & "-" & "|"
''                                    End If
''                        End Select
''                  Next j
''                  bd = bd & " " & ";"
'' '               End If
''                  i = i + 1
''            Loop
''
''            .AddTable fmt, hd, bd, , , False
''
''            .EndTable
''        .EndDoc
''    End With
''End Sub


''Private Sub GenerarDocF101Resumen(ByVal vp As SiiPrint.PreviewVSPrint, _
''                       ByVal grd As Control, _
''                       ByVal titulo As String, _
''                       ByVal rango_fecha As String)
''    Dim vpReport As VSPrinter, i As Long
''    'Dim celda_zon As String, celda_cli As String, celda_cat As String
''    Dim celda_cab As String, celda_tra As String, celda_des As String
''          Dim hd, j%, k%, fmt, bd As String, fila As Integer
''    Const COL_NUM = 0
''    Const COL_TIPO = 1
''    Const COL_CAMPO = 2
''    Const COL_TOTAL = 3
''    Set vpReport = vp.VSPrinter
''    With vpReport
''        .ShowGuides = gdShow
''        .Clear
''        .PhysicalPage = True
''        .MarginHeader = 800
''        .MarginTop = 800
''        .MarginBottom = 1200
''        .MarginFooter = 1200
''        .MarginLeft = 1200
''        .MarginRight = 800
''
''        .PaperSize = pprA4
''        .Orientation = orPortrait
''
''        .StartDoc
''        GeneraCabeceraGrafico vpReport, 100, titulo, rango_fecha
''        .FontName = "ARIAL"
''        .FontSize = 8
''        .Footer = "Fecha Impresión: " & Now & " USUARIO: " & _
''                  gobjMain.UsuarioActual.NombreUsuario & "||Pag: " & "%d"
''        .FontSize = 8
''        .FontBold = False
''        .CurrentY = 1200
''            .StartTable
''            .TableBorder = tbAll
''            .Paragraph = ""
''            i = 1
''                fila = 0
''                fmt = ">600|<1500|<1200|>1500"
''                hd = "#|TIPO|CAMPO|TOTAL"
''                Do While i <> grd.Rows
''                  grd.Select i, 1
''                  If grd.IsSubtotal(i) Then
''                  fila = fila + 1
''                  For j = 0 To 5
''                        grd.Select i, j
''                        Select Case j
''                              Case 0:
''                                          bd = bd & fila & "|"
''
''                              Case 2, 1:
''                                          bd = bd & grd.Text & "|"
''                              Case 5:
''                                    If grd.Text <> 0 Then
''                                          bd = bd & Format(grd.Text, "#,###0.00") & "|"
''                                    Else
''                                          bd = bd & "-" & "|"
''                                    End If
''                        End Select
''                  Next j
''                  bd = bd & " " & ";"
''               End If
''                  i = i + 1
''            Loop
''
''            .AddTable fmt, hd, bd, , , False
''
''            .EndTable
''        .EndDoc
''    End With
''End Sub
''
''
''
''
