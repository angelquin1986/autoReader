Attribute VB_Name = "modSiiComun"
Option Explicit

'Modulo de procedimientos comunes entre todos los proyectos de Sii3
'#Const DEBUGMODE = 1
Public Function CogeSoloCamino(ByVal s As String) As String
    Dim i As Integer
    'Busca la ultima posicion de "\"
    For i = Len(s) To 1 Step -1
        If Mid$(s, i, 1) = "\" Then Exit For
    Next i
    s = Left$(s, i)
    CogeSoloCamino = s
End Function

Public Function CogeSoloNombre(ByVal s As String) As String
    Dim i As Integer
    'Busca la ultima posicion de "\"
    For i = Len(s) To 1 Step -1
        If Mid$(s, i, 1) = "\" Then Exit For
    Next i
    s = Right$(s, Len(s) - i)
    CogeSoloNombre = s
End Function

Public Function QuitaExtension(ByVal s As String) As String
    Dim i As Integer
    'Busca la ultima posicion de "."
    For i = Len(s) To 1 Step -1
        If Mid$(s, i, 1) = "." Then Exit For
    Next i
    
    If i > 0 Then s = Left$(s, i - 1)
    QuitaExtension = s
End Function


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

Public Function Redondear(ByVal v As Currency, ByVal digitos As Integer) As Currency
'v :        Valor para redondear
'digitos :  Numeros de digitos a redondear
'           *** Decimales son negativos.
'               ejm. para redondear hasta 0.01 (2 decimales) envie -2
'               ejm. para redondear hasta 100 envie 2
    
'***Redondeo de Round es 'Redondeo bancario'
'   Es decir:   2.5 ==> 2, 3.5 ==> 4, 4.5 ==> 4, 5.5 ==> 6  ...
    Dim a As Long, r As Currency
    r = 10 ^ digitos
    If r > 1 Then
        a = Round((v / r), 0)
        v = a * r
    Else
        v = Round(v, Abs(digitos))
    End If
    Redondear = v
End Function

'*** MAKOTO 29/ene/01 Agregado
'Para convertir de otro tipo a Currency sin que se de error de conversión
Public Function MiCCur(ByVal v As Variant) As Currency
    If IsNumeric(v) Then MiCCur = CCur(v)
End Function

Public Sub DispErr(Optional msg As String)
    Dim s As String
    If Len(msg) > 0 Then
        s = msg
        If Err.Number <> 0 And Err.Number < ERRNUM Then
            s = s & vbCr & Err.Number
        End If
    Else
        If Err.Number < ERRNUM Then s = Err.Number
    End If

    If Len(s) > 0 Then s = s & vbCr & vbCr
    If Err.Number = ERR_DESBORDA Then           '*** MAKOTO 15/feb/01 mod.
        s = s & MSGERR_DESBORDA     'Si es Desbordamiento, sacar mensaje un poco más entendible
    Else
        s = s & Err.Description
    End If
#If DEBUGMODE Then
    If Len(Err.Source) > 0 Then
        s = s & vbCr & vbCr
        s = s & "(" & Err.Source & ")"
    End If
#End If
    MsgBox s, vbInformation
        End Sub

Public Sub MoverCampo(frm As Form, KeyCode As Integer, Shift As Integer, UnloadParaCerrar As Boolean)
    Dim i As Integer, st As Integer, ed As Integer, act As Control
    Dim key_close As Integer, t As String
    
    'Si aplastó ESC
    If KeyCode = vbKeyEscape Then
        'Cierra u oculta la pantalla
        If UnloadParaCerrar Then Unload frm Else frm.Hide
        Exit Sub
    End If
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Set act = frm.ActiveControl
    If act Is Nothing Then Exit Sub
    t = UCase$(TypeName(act))
    If t = "VSFLEXGRID" Then Exit Sub
    If t = "ENCABEZADO" Then Exit Sub
    If t = "ASIENTO" Then Exit Sub
    If t = "IVGN" Then Exit Sub
    If t = "IVGNBUSQUEDA" Then Exit Sub
    If t = "IVREC" Then Exit Sub
    If t = "PCDOC" Then Exit Sub
    If t = "TSDOC" Then Exit Sub
    If t = "TSRET" Then Exit Sub            '*** MAKOTO 08/feb/01 Agregado
    KeyCode = 0
    
    'Si no está presionada la tecla 'Shift', se mueve para adelante
    If Not (Shift And vbShiftMask) > 0 Then
        st = act.TabIndex + 1
        ed = frm.Controls.count - 1
        If Not MoverCampoSub(st, ed, frm) Then
            st = 0
            ed = act.TabIndex - 1
            MoverCampoSub st, ed, frm
        End If
    
    'Si está presionada la tecla 'Shift' se mueve para atrás
    Else
        st = act.TabIndex - 1
        ed = 0
        If Not MoverCampoSub(st, ed, frm) Then
            st = frm.Controls.count - 1
            ed = act.TabIndex + 1
            MoverCampoSub st, ed, frm
        End If
    End If
'Debug.Print "Movido a " & frm.ActiveControl.Name
End Sub

Private Function MoverCampoSub(st As Integer, ed As Integer, frm As Form) As Boolean
    Dim i As Integer, c As Control
    On Error GoTo ErrTrap
    
    If st = ed Then Exit Function
    
    MoverCampoSub = True
    For i = st To ed Step Sgn(ed - st)
        For Each c In frm.Controls
            'Si no tienen TabIndex o no deben tener enfoque
            If (TypeName(c) = "CommonDialog") Or _
               (TypeName(c) = "CrystalReport") Or _
               (TypeName(c) = "Label") Or _
               (TypeName(c) = "Menu") Or _
               (TypeName(c) = "Line") Or _
               (TypeName(c) = "Timer") Or _
               (TypeName(c) = "Frame") Or _
               (TypeName(c) = "Data") Or _
               (TypeName(c) = "ImageList") Or _
               (TypeName(c) = "Image") Or _
               (TypeName(c) = "Toolbar") Or _
               (TypeName(c) = "SelFolder") Or _
               (TypeName(c) = "Picture") Then
                'No hace nada
            Else
                If c.TabIndex = i Then
                    If c.Enabled And _
                       c.Visible And _
                       c.TabStop And _
                       c.Left >= 0 And c.Top >= 0 Then
                        c.SetFocus
                        Exit Function
                    End If
                    Exit For
                End If
            End If
        Next c
    Next i
    
    MoverCampoSub = False
    Exit Function
ErrTrap:
    DispErr c.Name
'    Resume Next
    Exit Function
End Function

Public Sub ImpideSonidoEnter(frm As Form, KeyAscii As Integer)
    Dim t As String
    
    'Si no hay ningun control activo no hace nada
    If (frm.ActiveControl Is Nothing) Then Exit Sub
    
    'Controles que no se debe ignorar Enter
    t = UCase$(TypeName(frm.ActiveControl))
    If t = "VSFLEXGRID" Then Exit Sub
    If t = "ENCABEZADO" Then Exit Sub
    If t = "ASIENTO" Then Exit Sub
    If t = "IVGN" Then Exit Sub
    If t = "IVGNBUSQUEDA" Then Exit Sub
    If t = "IVREC" Then Exit Sub
    If t = "PCDOC" Then Exit Sub
    If t = "TSDOC" Then Exit Sub
    If t = "TSCOBROPAGO" Then Exit Sub
    If t = "TSRET" Then Exit Sub            '*** MAKOTO 12/feb/01 Agregado
        
    'Ignora la tecla 'Enter' para que no se genere el sonido
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub


'Detecta y devuelve simbolo de punto decimal
Public Function SimboloDecimal() As String
    Dim s As String
    
    'Detecta simbolo de decimales
    s = Format(1.5, "0.00")
    s = Mid$(s, 2, 1)
    SimboloDecimal = s
End Function

'Detecta y devuelve simbolo separación de miles
Public Function SimboloMiles() As String
    Dim s As String
    
    'Detecta simbolo de decimales
    s = Format(9999, "0,0")
    s = Mid$(s, 2, 1)
    SimboloMiles = s
End Function



'*** MAKOTO 03/oct/2000 Agregado
'Valida las teclas
'En columna de tipo numérico solo acepta numéricos
Public Sub ValidarTeclaFlexGrid( _
            ByVal grd As Object, _
            ByVal Row As Long, _
            ByVal col As Long, _
            ByRef KeyAscii As Integer, _
            Optional ByVal NoNegativo As Boolean)
    Dim sd As String
    
    'Detecta simbolo de decimales
    sd = SimboloDecimal
            
    Select Case grd.ColDataType(col)
    Case flexDTCurrency, flexDTSingle, flexDTDouble
        If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
            (KeyAscii <> vbKeyBack) And _
            (KeyAscii <> Asc(sd)) And _
            (KeyAscii <> Asc("-") Or (NoNegativo)) And _
            (KeyAscii <> vbKeyReturn) And _
            (KeyAscii <> 22) Then               '22 = CTRL+v (CTRL+c es automático)
            KeyAscii = 0
        End If
    Case flexDTLong, flexDTShort
        If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
            (KeyAscii <> vbKeyBack) And _
            (KeyAscii <> vbKeyReturn) And _
            (KeyAscii <> Asc("-") Or (NoNegativo)) And _
            (KeyAscii <> 22) Then               '22 = CTRL+v (CTRL+c es automático)
            KeyAscii = 0
        End If
    End Select
End Sub

Public Sub DebugTimer(msg As String, inicializa As Boolean)
'msg : Cualquier mensaje

    Static Inicio As Single, anterior As Single, t As Single
    
    t = Timer
    If inicializa Then
        Inicio = t
        anterior = Inicio
        Debug.Print "-DebugTimer:" & msg & " Inicialización."
    Else
        Debug.Print "-DebugTimer:" & msg & _
                ": " & Format((t - anterior), "00.0000") & vbTab & _
                " suma=" & Format((t - Inicio), "00.0000")
    End If
    anterior = t
End Sub

Public Sub DebugPrintCols(ByVal rs As Object)
    Dim i As Long
    For i = 0 To rs.Fields.count
        Debug.Print i, rs.Fields.item(i).Name
    Next i
End Sub

Public Function RedondeaMoneda(ByVal v As Currency, ByVal fmt As String) As Currency
    Dim numd As Integer, i As Integer, a As Single
    'Obtiene número de digitos decimales fraccionales
    i = InStrRev(fmt, ".")
    numd = 0
    If i > 0 Then numd = Len(fmt) - i
    If numd >= 0 Then a = 5 / (10 ^ (numd + 1))
    If numd = 0 Then
        v = Fix(v + a)
    Else
'        If Abs(v - Fix(v)) >= a Then v = v + a / 5
'        v = Round(v, numd)
    End If
    RedondeaMoneda = v
End Function

'----- Verifica si existe un archivo
Public Function ExisteArchivo(archi As String) As Boolean
    Dim n As Integer
    On Error GoTo ErrTrap   'Necesario siempre
    
    n = FreeFile
    Open archi For Input As #n
    Close n
    ExisteArchivo = True
    Exit Function
ErrTrap:
    Exit Function
End Function

Public Sub LimpiaColeccion(c As Collection)
    Dim i As Long
    For i = c.count To 1 Step -1
        c.Remove i
    Next i
End Sub

'*** MAKOTO 01/oct/2000 Agregado
'Ajustar automáticamente ancho de columnas y alto de filas
'   grd : Objeto VsFlexGrid para ajustar
'   Row : Indice de fila que el usuario cambio su alto.(Envia de grd_AfterUserResize)
'         Debe enviar -1 si no está llamando de grd_AfterUserResize
'   Col : Indice de columna que el usuario cambio su ancho.(Envia de grd_AfterUserResize)
'         Debe enviar -1 si no está llamando de grd_AfterUserResize
'   colWidthMax : Ancho maximo de columnas.
'                   (Esto se aplica solo temporalmente y no afecta a la propiedad 'ColWidthMax' de la grilla)
Public Sub AjustarAutoSize( _
                ByVal grd As Object, _
                ByVal Row As Long, _
                ByVal col As Long, _
                Optional ByVal colWidthMax)
    On Error GoTo ErrTrap
    
    Screen.MousePointer = vbHourglass
    If IsMissing(colWidthMax) Then colWidthMax = 4500
    
    With grd
        .WordWrap = True
        .colWidthMax = colWidthMax
        If col < 0 Then
            .AutoSizeMode = flexAutoSizeColWidth
            .AutoSize 0, .Cols - 1
        End If
        If Row < 0 Then
            .AutoSizeMode = flexAutoSizeRowHeight
            .AutoSize 0, .Cols - 1
        End If
    End With
    Screen.MousePointer = vbNormal
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbNormal
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Exit Sub
End Sub

Public Sub AsignarTituloAColKey(ByVal grd As Object)
    Dim i As Long
    With grd
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
        Next i
    End With
End Sub

'*** MAKOTO 12/dic/00
'Cada vez que llama ésta función, cambia color de fondo y títulos de VSFlexGrid
'Debe llamar desde GotFocus y LostFocus ambos
Public Sub FlexGridGotFocusColor( _
                ByVal grd As Control)
    With grd
        If .Editable Then
            .ForeColorFixed = grd.ForeColorFixed Xor &HCC0101
            .BackColorBkg = .BackColorBkg Xor &H999999
        End If
    End With
End Sub

'*** MAKOTO 29/ene/01 Agregado.
'Asegura que no haya diferencia entre el contenido de una
'celda y lo que está visualizando
Public Sub FlexGridRedondear( _
                ByVal grd As Object, _
                ByVal Row As Long, _
                ByVal col As Long)
    With grd
        .TextMatrix(Row, col) = .Cell(flexcpTextDisplay, Row, col)
    End With
End Sub

'jeaa 15/03/2005

Function RegGet(regKey, Optional DefaultValue = "")
    On Error GoTo ErrTrap
    Set Regedit = CreateObject("WScript.Shell")
    RegGet = Regedit.RegRead(regKey)
    Set Regedit = Nothing
    Exit Function
ErrTrap:
    RegGet = DefaultValue
    Set Regedit = Nothing
End Function

Public Sub RecuperaSelec(ByVal Key As String, lst As ListBox, Optional s As String)
Dim Vector As Variant
Dim i As Integer, j As Integer, Selec As Integer
'Dim S As String
    If s <> "_VACIO_" Then
        Vector = Split(s, ",")
         Selec = UBound(Vector, 1)
         For i = 0 To Selec
            For j = 0 To lst.ListCount - 1
                If Vector(i) = Left(lst.List(j), lst.ItemData(j)) Then
                    lst.Selected(j) = True
                End If
            Next j
         Next i
    End If
End Sub

Public Sub RecuperaSelecparaIN(ByVal Key As String, lst As ListBox, Optional s As String)
Dim Vector As Variant
Dim i As Integer, j As Integer, Selec As Integer
'Dim S As String
    If s <> "_VACIO_" And Len(s) <> 1 Then
        Vector = Split(s, ",")
         Selec = UBound(Vector, 1)
         
         For i = 0 To Selec
            For j = 0 To lst.ListCount - 1
                If Mid$(Vector(i), 2, Len(Vector(i)) - 2) = Left(lst.List(j), lst.ItemData(j)) Then
                    lst.Selected(j) = True
                End If
            Next j
         Next i
    End If
End Sub

'*** MAKOTO 29/ene/01 Agregado
'Para convertir de otro tipo a Double sin que se de error de conversión
Public Function MiCDbl(ByVal v As Variant) As Double
    If IsNumeric(v) Then MiCDbl = CDbl(v)
End Function

Public Sub ValidarTeclaNumeros( _
            ByVal t As Object, _
            ByRef KeyAscii As Integer, _
            Optional ByVal NoNegativo As Boolean)
    Dim sd As String
    sd = SimboloDecimal
        If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
            (KeyAscii <> vbKeyBack) And _
            (KeyAscii <> Asc(sd)) And _
            (KeyAscii <> Asc("-") Or (NoNegativo)) And _
            (KeyAscii <> vbKeyReturn) And _
            (KeyAscii <> 22) Then               '22 = CTRL+v (CTRL+c es automático)
            KeyAscii = 0
        End If
    
End Sub

'AUC 11/2011
Public Sub RecuperaSelecRol(ByVal Key As String, grd As VSFlexGrid, s As String, col As Long)
Dim Vector As Variant
Dim i As Integer, j As Integer, Selec As Integer
    If s <> "_VACIO_" Then
        Vector = Split(s, ",")
         Selec = UBound(Vector, 1)
         For i = 0 To Selec
            For j = 0 To grd.Rows - 1
            If Vector(i) = grd.TextMatrix(j, 1) Then
                 grd.TextMatrix(j, col) = -1
            End If
            Next j
         Next i
    End If
End Sub


