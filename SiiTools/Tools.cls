VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, _
                                        ByVal hWndNewParent As Long) As Long


'Para reprocesamiento de costos
Public Sub ShowReprocCosto(ByVal objSiiMain As SiiMain, _
                       Optional ByVal hWndPadre As Long)
    If objSiiMain.EmpresaActual Is Nothing Then
        Err.Raise ERR_INVALIDO, "Tools.ReprocCosto", _
            "No hay una empresa abierta."
    End If

    If hWndPadre <> 0 Then
        SetParent frmPrecios.hwnd, hWndPadre
    End If
    frmReprocCosto.Inicio objSiiMain
End Sub



'Para la actualización de precios
Public Sub ShowPrecios(ByVal objSiiMain As SiiMain, _
                       Optional ByVal hWndPadre As Long)
    If objSiiMain.EmpresaActual Is Nothing Then
        Err.Raise ERR_INVALIDO, "Tools.ShowPrecios", _
            "No hay una empresa abierta."
    End If
    
    If hWndPadre <> 0 Then
        SetParent frmPrecios.hwnd, hWndPadre
    End If
    frmPrecios.Inicio objSiiMain
End Sub

'Para la regeneración de asientos
Public Sub ShowAsiento(ByVal objSiiMain As SiiMain, _
                       Optional ByVal hWndPadre As Long)
    If objSiiMain.EmpresaActual Is Nothing Then
        Err.Raise ERR_INVALIDO, "Tools.ShowPrecios", _
            "No hay una empresa abierta."
    End If

End Sub

'Para la actualización de %IVA
Public Sub ShowIVA(ByVal objSiiMain As SiiMain, _
                       Optional ByVal hWndPadre As Long)
    If objSiiMain.EmpresaActual Is Nothing Then
        Err.Raise ERR_INVALIDO, "Tools.ShowPrecios", _
            "No hay una empresa abierta."
    End If

End Sub





