Attribute VB_Name = "ModuloRol"
Option Explicit

Public Sub MainRoles()
    Dim code As String
    On Error GoTo ErrTrap


    Set gobjRol = PreparaSiiMain
    gobjRol.Inicializar
    
'    If Not frmLoginRol.Inicio Then End
'
'    Unload frmLoginRol

    'frmGeneraAsientoRol.Show
    
    'Obtiene codigo de la ultima empresa
    code = gobjRol.EmpresaAnterior
    'Si no puede recuperar, selecciona
    If Len(code) = 0 Then
        frmSelecEmpRol.Show vbModal
    'Si recupera la empresa anterior, abre la misma
    Else
        'Si no la puede abrir, hace seleccionar
        If Not AbrirEmpresaSii(code, False) Then
            frmSelecEmpRol.Show vbModal
        End If
    End If
    
    
    
    Exit Sub
    
ErrTrap:
        If Err.Number = ERR_NOREGINFO Then
        MsgBox "El programa se inicia por primera vez." & vbCr & _
               "Llene las configuraciones iniciales del sistema."
    Else
        DispErr
    End If
'    Unload frmLoginRol
    End
    Exit Sub
End Sub

'Recibe codigo de empresa y la abre
Public Function AbrirEmpresaSii(ByVal cod As String, ByVal mensaje As Boolean) As Boolean
    Dim emp As Object
    On Error GoTo ErrTrap
    
    AbrirEmpresaSii = False
    
    MensajeStatus "Está abriendo la empresa ...", vbHourglass
    Set emp = gobjRol.RecuperaEmpresaDesdeSII(cod)
    
    If Not (emp Is Nothing) Then
        If Not (gobjRol.EmpresaActual Is Nothing) Then
            gobjRol.EmpresaActual.Cerrar
        End If
        
        'Abre la base de datos de la empresa
        emp.Abrir
        AbrirEmpresaSii = True
    ElseIf mensaje Then
        MensajeStatus "", 0
        MsgBox "No se puede abrir la empresa '" & cod & "'."
    End If
    Set emp = Nothing
    
    'frmMain.CambiaCaption       'Actualiza la Caption
    MensajeStatus "", 0
    Exit Function
ErrTrap:
    MensajeStatus "", 0
    If mensaje Then DispErr
    Exit Function
End Function

Public Sub SeleccionaComboItem(cbo As ComboBox, cod As String)
    Dim i As Integer
    
    cbo.ListIndex = -1
    For i = 0 To cbo.ListCount - 1
        If cbo.List(i) = cod Then
            cbo.ListIndex = i
        End If
    Next i
End Sub

Public Sub FlexAlterarCheck(ByVal grd As VSFlexGrid)
    Dim i As Long
    With grd
        For i = 0 To .SelectedRows - 1
            If .ValueMatrix(.SelectedRow(i), .ColIndex("Check")) <> 0 Then
                .Cell(flexcpChecked, .SelectedRow(i), .ColIndex("Check")) = flexUnchecked
            Else
                .Cell(flexcpChecked, .SelectedRow(i), .ColIndex("Check")) = flexChecked
            End If
        Next i
    End With
End Sub



'***Angel. 05/ene/2003.
'***Para trabajar con Dll de Sii en tiempo de ejecucion y no por referencia
Public Function PreparaSiiMain() As Object
    Dim objMain As Object
    Set PreparaSiiMain = Nothing
    
    'Crea objeto SiiMain
    On Error Resume Next
    
    Set objMain = CreateObject("rol_pagosa.RolMain")
    If Err.Number = 0 Then
        Err.Clear
    Else
        Err.Clear
        Set objMain = CreateObject("rolDLLA.rolMain")
        If Err.Number = 0 Then
            Err.Clear
        Else
            MsgBox "No está instalado el sistema SII en ninguna de sus versiones", vbExclamation
            Exit Function
        End If
    End If
    Set PreparaSiiMain = objMain
End Function
Public Sub AbrirEmpresaReloj(ByVal cod As String, ByVal mensaje As Boolean, ByRef objReloj As Object)
    Dim emp As Object, i As Byte
    On Error GoTo ErrTrap
    
    
'    AbrirEmpresaReloj = False
    
    MensajeStatus "Está abriendo la empresa ...", vbHourglass
    Set emp = objReloj.RecuperarEmpresa(cod)
    
    If Not (emp Is Nothing) Then
        If Not (objReloj.EmpresaActual Is Nothing) Then
            objReloj.EmpresaActual.Cerrar
        End If
        
        'Abre la base de datos de la empresa
        emp.Abrir
        'AbrirEmpresaSii4 = True
        
    ElseIf mensaje Then
        MensajeStatus "", 0
        MsgBox "No se puede abrir la empresa '" & cod & "'."
    End If
    Set emp = Nothing
           MensajeStatus "", 0
    Exit Sub
ErrTrap:
    MensajeStatus "", 0
    If mensaje Then DispErr
    Exit Sub
End Sub

