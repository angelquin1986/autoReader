Attribute VB_Name = "ExportaTexto"
Option Explicit

Public Sub ExportaTxt(ByVal grd As Control, _
                      ByRef archi As String)

    Dim file As String, NumFile As Integer, fila
    Dim r As Long, c As Long
    NumFile = FreeFile
    file = CargaNombreArchivo(archi)
    If file = "" Then
        archi = file
        Exit Sub
    End If
    
    Open file For Output Access Write As #NumFile
    With grd
        If Mid(file, InStr(file, ".") + 1, 3) = "txt" Then
            For r = 0 To .Rows - 1
                If .RowHidden(r) = False Then ' Filas  Ocultas
                    fila = ""
                    For c = 1 To .Cols - 1
                        If .ColHidden(c) = False Then
                            fila = fila & """" & .TextMatrix(r, c) & ""","
                        End If
                    Next c
                    fila = Left(fila, Len(fila) - 1)
                    Print #NumFile, fila
                End If
            Next r
        Else
            For r = 0 To .Rows - 1
                If .RowHidden(r) = False Then ' Filas  Ocultas
                    fila = ""
                    For c = 1 To .Cols - 1
                        If .ColHidden(c) = False Then
                            fila = fila & .TextMatrix(r, c) & ","
                        End If
                    Next c
                    fila = Left(fila, Len(fila) - 1)
                    Print #NumFile, fila
                End If
            Next r
        End If
    End With
    Close NumFile
    archi = file
End Sub

Public Sub ExportaTxtConTab(ByVal grd As Control, _
                            ByRef archi As String)

    Dim file As String, NumFile As Integer, fila
    Dim r As Long, c As Long
    NumFile = FreeFile
    file = CargaNombreArchivo(archi)
    If file = "" Then
        archi = file
        Exit Sub
    End If
    
    Open file For Output Access Write As #NumFile
    With grd
        If Mid(file, InStr(file, ".") + 1, 3) = "txt" Then
            For r = 0 To .Rows - 1
                If .RowHidden(r) = False Then ' Filas  Ocultas
                    fila = ""
                    For c = 1 To .Cols - 1
                        If .ColHidden(c) = False Then
                            fila = fila & """" & .TextMatrix(r, c) & """" & vbTab
                        End If
                    Next c
                    fila = Left(fila, Len(fila) - 1)
                    Print #NumFile, fila
                End If
            Next r
        Else
            For r = 0 To .Rows - 1
                If .RowHidden(r) = False Then ' Filas  Ocultas
                    fila = ""
                    For c = 1 To .Cols - 1
                        If .ColHidden(c) = False Then
                            fila = fila & .TextMatrix(r, c) & ","
                        End If
                    Next c
                    fila = Left(fila, Len(fila) - 1)
                    Print #NumFile, fila
                End If
            Next r
        End If
    End With
    Close NumFile
    archi = file
End Sub

Private Function CargaNombreArchivo(ByVal file As String) As String
    On Error GoTo ErrTrap
    With frmMain.dlg1
        .InitDir = App.Path
        .DialogTitle = "Guardar Archivo"
        .CancelError = True
        .Filter = "Archivo de Texto|*.txt;|Texto (Separado por coma)|*.csv"
        .DefaultExt = "txt"
        .flags = cdlOFNCreatePrompt Or cdlOFNOverwritePrompt Or cdlOFNHideReadOnly
        If Len(file) > 0 Then .filename = file
        .ShowSave
        CargaNombreArchivo = .filename
    End With
    Exit Function
ErrTrap:
    Exit Function
End Function
