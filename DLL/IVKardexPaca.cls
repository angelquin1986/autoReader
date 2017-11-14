VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVKardexPaca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROP
    Id As Long

    IdBodega As Long
    CodBodega As String
    Cantidad As Currency
    Nota As String
    Orden As Integer
    TransidDet As Long
    objGNComprobante As GNComprobante           'Referencia al objeto padre
End Type

Private mcolConsumoDetalle As Collection      'Coleccion de detalles de familia
Private mcolImportDetalle As Collection      'Coleccion de detalles de familia
Private mcolIVSerie As Collection
Private mProps As T_PROP

Private Sub AseguraBodega()
    Dim sql As String, rs As Recordset
    
    sql = "SELECT IdBodega FROM IVExist WHERE IdInventario=" & _
                mProps.IdInventario & " AND IdBodega=" & mProps.IdBodega
    
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
    'Si no existe el registro correspondiente, lo inserta
    If rs.EOF Then
        sql = "INSERT INTO IVExist (IdInventario,IdBodega,Exist) " & _
              "VALUES(" & mProps.IdInventario & "," & mProps.IdBodega & "," & 0 & ")"
        mProps.objGNComprobante.Empresa.Execute sql, True
    End If
    rs.Close
    Set rs = Nothing
End Sub

Public Sub VerificaDatos()
    Dim cant As Currency
    Dim item As IVInventario
    Dim pc As PCProvCli, nunPCGrupo As Integer, cadPrecios As String, idPcGrupo As Long
    Dim nprecio As Integer
    Dim ivs As IVNumSerie
    Dim bandSerie As Boolean
    Dim rs As Recordset
    Dim i As Long
    Dim ContNumSerie As Long
'    Set item = mProps.objGNComprobante.Empresa.RecuperaIVInventario(mProps.IdInventario)
       
           
     If Not mProps.objGNComprobante.GNTrans.MostrarSubItemsImp Then  'AUC cambiado para arreglar la secuencia
        Set item = mProps.objGNComprobante.Empresa.RecuperaIVInventario(mProps.CodInventario)
        If item.Tipo = Preparacion Then
            'mProps.Cantidad = -1
            mProps.Cantidad = Me.Cantidad
        End If
    End If
       'La cantidad no puede ser 0.
    Set item = mProps.objGNComprobante.Empresa.RecuperaIVInventario(mProps.CodInventario)
    If item.Tipo <> rubro Then
        If mProps.Cantidad = 0 Then
           Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", _
              "La cantidad no puede ser 0. (" & mProps.CodInventario & ")"
        End If
    End If
    If mProps.objGNComprobante.GNTrans.IVGeneraNumSerie Then
       If item.bandSerie Then
            If mProps.objGNComprobante.CountIVKNumSerie = 0 Then
                     Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", _
                    "Debe Ingresar Numero de serie para esta transacción"
            End If
                Set rs = mProps.objGNComprobante.Empresa.RecuperaIdSerieItem(mProps.IdInventario)
                Do While Not rs.EOF
                    For i = 1 To mProps.objGNComprobante.CountIVKNumSerie
                        If mProps.objGNComprobante.GNTrans.IVTipoTrans = "I" Then
                           If rs!IdSerie = mProps.objGNComprobante.IVKNumSerie(i).IdSerie And mProps.Orden = mProps.objGNComprobante.IVKNumSerie(i).OrdenIvkardex Then
                                bandSerie = True
                                ContNumSerie = ContNumSerie + 1
                           End If
                        Else 'Para transf. y egresos
                            If mProps.objGNComprobante.IVKNumSerie(i).Cantidad < 0 Then
                                If rs!IdSerie = mProps.objGNComprobante.IVKNumSerie(i).IdSerie Then
                                    bandSerie = True
                                    ContNumSerie = ContNumSerie + 1
                                End If
                            End If
                        End If
                    Next
                    rs.MoveNext
                Loop
                If Not bandSerie Then
                    Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", _
                    "Debe Ingresar Numero de serie para este Item " & mProps.CodInventario
                End If
                If mProps.objGNComprobante.GNTrans.IVTipoTrans = "I" Then
                    If ContNumSerie <> mProps.Cantidad Then
                        Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", _
                        "La cantidad de Items no coincide con el Numero de Series asignadas" & Chr(13) & "Item: " & mProps.CodInventario
                    End If
                Else
                   ' If mProps.objGNComprobante.IVKNumSerie(i).Cantidad < 0 Then
                        If ContNumSerie <> Abs(mProps.Cantidad) Then
                            Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", _
                            "La cantidad de Items no coincide con el Numero de Series asignadas" & Chr(13) & "Item: " & mProps.CodInventario
                        End If
                    'End If
                End If
            Set ivs = Nothing
            Set rs = Nothing
        End If
    End If
       
       'La cantidad no puede ser 0.
       If mProps.Cantidad = 0 Then
            Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", _
                "La cantidad no puede ser 0. (" & mProps.CodInventario & ")"
         End If
'    End If
    
    If mProps.IdBodega = 0 Then
        Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", _
            "Debe seleccionar una bodega. (" & mProps.CodInventario & ")"
    End If

    If mProps.IdInventario = 0 Then
        Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", _
            "Debe seleccionar un item."
    End If
    
    '*** MAKOTO 07/feb/01 Agregado
    'Si la transacción está configurado que NO permita grabar con costo 0
    If mProps.objGNComprobante.GNTrans.IVNoPermitirCostoCero And Not item.BandServicio Then
        If mProps.CostoRealTotal = 0 Then
            Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", _
                "En ésta transacción no se puede grabar con costo 0. (" & mProps.CodInventario & ")"
        End If
    End If
    '*** Oliver 15/07/2003 Agregado
    'Si la transaccion esta configurada para controlar que no graben con percios 0 en los items
    
    If mProps.objGNComprobante.GNTrans.CodPantalla = "IVPVTM" Then
    
        If mProps.objGNComprobante.GNTrans.IVControlPrecio0 And mProps.IdPadreSub = 0 And Not item.BandRepGastos Then
            If mProps.PrecioTotal = 0 Then
                Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", _
                    "En ésta transacción no se puede grabar con precio 0. (" & mProps.CodInventario & ")"

            End If
        End If
    
    
    ElseIf mProps.objGNComprobante.GNTrans.CodPantalla <> "IVPVTS" And item.CodGrupo(1) <> "00004" Then
        If mProps.objGNComprobante.GNTrans.IVControlPrecio0 Then
            If mProps.PrecioRealTotal = 0 Then
                Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", _
                    "En ésta transacción no se puede grabar con precio 0. (" & mProps.CodInventario & ")"
            End If
        End If
    End If
    'Modificado  DIEGO 09/06/2001
    Dim cantOtras As Currency, msg As String, CantLimite As Currency, Saldo As Currency
    Dim fmt As String
    CantLimite = 0
    cantOtras = 0

            'control para HORMIPISOS
    If mProps.objGNComprobante.GNTrans.IVControlaItemServicio Then
        If item.BandServicio Then Exit Sub
    End If
    
    If mProps.objGNComprobante.GNTrans.IVControlaSoloCantidad Then
        If Not mProps.objGNComprobante.GNTrans.ControlCantidadxFamilia Then
            If mProps.objGNComprobante.VerificaItemConFuenteSoloCantidadxIVGrupo(mProps.objGNComprobante.CodTrans, mProps.objGNComprobante.NumTrans, mProps.objGNComprobante.IdTransFuente, Me.CodInventario, CantLimite, cantOtras) Then
                If Not mProps.objGNComprobante.GNTrans.IvCambioColorGris Then
                    Saldo = (CantLimite) - cantOtras
    
                    If Abs(Me.Cantidad) > Abs(Saldo) Then
                        fmt = mProps.objGNComprobante.Empresa.GNOpcion.FormatoCantidad
                        msg = "Ha intentado registrar la cantidad mayor al límite " & _
                              "establecido en la transacción fuente" & _
                              vbCr & vbCr & _
                              "    Código de item: " & Me.CodInventario & vbCr & _
                              "    Cantidad límite: " & Format(Abs(CantLimite), fmt) & _
                              "    Cantidad utilizada: " & Format(Abs(cantOtras), fmt) & _
                              "    Saldo: " & Format(Abs(Saldo), fmt)
                        Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", msg
                    End If
                Else
                    'para cambiar de color a gris solo con clave de supervisor
                    Saldo = (CantLimite) - cantOtras
    
                    If Abs(Me.Cantidad) > Abs(Saldo) Then
                        fmt = mProps.objGNComprobante.Empresa.GNOpcion.FormatoCantidad
                        msg = "Va ha registrar la cantidad mayor al límite " & _
                              "establecido en la transacción fuente" & _
                              vbCr & vbCr & _
                              "    Código de item: " & Me.CodInventario & vbCr & _
                              "    Cantidad límite: " & Format(Abs(CantLimite), fmt) & _
                              "    Cantidad utilizada: " & Format(Abs(cantOtras), fmt) & _
                              "    Saldo: " & Format(Abs(Saldo), fmt)
                        Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", msg
                    End If
                End If
            End If
        Else
                If mProps.objGNComprobante.VerificaItemConFuenteSoloCantidadFamilia(Me.CodInventario, CantLimite) Then
                If Not mProps.objGNComprobante.GNTrans.IvCambioColorGris Then
                        Saldo = (CantLimite)
                        If Abs(Me.Cantidad) > Abs(Saldo) Then
                            fmt = mProps.objGNComprobante.Empresa.GNOpcion.FormatoCantidad
                            msg = "Ha intentado registrar la cantidad mayor al límite " & _
                                  "establecido en la transacción fuente" & _
                                  vbCr & vbCr & _
                                  "    Código de item: " & Me.CodInventario & vbCr & _
                                  "    Cantidad límite: " & Format(Abs(CantLimite), fmt) & _
                                  "    Cantidad utilizada: " & Format(Abs(cantOtras), fmt) & _
                                  "    Saldo: " & Format(Abs(Saldo), fmt)
                            Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", msg
                        End If
                Else
                        If Abs(Me.Cantidad) > Abs(Saldo) Then
                            fmt = mProps.objGNComprobante.Empresa.GNOpcion.FormatoCantidad
                            msg = "Ha intentado registrar la cantidad mayor al límite " & _
                                  "establecido en la transacción fuente" & _
                                  vbCr & vbCr & _
                                  "    Código de item: " & Me.CodInventario & vbCr & _
                                  "    Cantidad límite: " & Format(Abs(CantLimite), fmt) & _
                                  "    Cantidad utilizada: " & Format(Abs(cantOtras), fmt) & _
                                  "    Saldo: " & Format(Abs(Saldo), fmt)
                            'Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", msg
                        End If
                
                End If
            End If
        End If
    End If
    
    
    'Controla  que  sean los mismos Items de la transaccion fuente
    If mProps.objGNComprobante.GNTrans.IVVerificaItemsFuente Then
        If mProps.objGNComprobante.IdTransFuente <> 0 Then
            'Obtiene cantidad del mismo item en otras filas del mismo comprobante
            cantOtras = SumaItemRepetido '+ Me.Cantidad
            'Recupera transaccion Fuente y  Verifica Limite de cantidad
            
            If mProps.objGNComprobante.VerificaItemConFuente(Me.CodInventario, CantLimite) = False Then
                Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", _
                "Código de Item: " & mProps.CodInventario & " no existe en la transaccion fuente"
            Else
                If CantLimite = 0 Then
                    Saldo = 0
                Else
                    Saldo = CantLimite
                'solo si las transacciones son diferentes Egreso /Ingreso
                '                                         Ingreso / Egreso
                    
                    If mProps.objGNComprobante.GNTrans.IVTipoTrans = "I" Then
                        If Saldo > 0 Then Saldo = 0
                    Else
                        If Saldo < 0 Then Saldo = 0
                    End If
                End If
                Saldo = Abs(Saldo) - cantOtras
                If mProps.objGNComprobante.GNTrans.IVItemDevolucion Then
                    If mProps.objGNComprobante.GNTrans.CodLineaDevolucion = mProps.CodInventario Then
                        Saldo = 1
                    End If
                End If

                If Me.Cantidad > Saldo Then
                    fmt = mProps.objGNComprobante.Empresa.GNOpcion.FormatoCantidad
                    msg = "Ha intentado registrar la cantidad mayor al límite " & _
                          "establecido en la transacción fuente" & _
                          vbCr & vbCr & _
                          "    Código de item: " & Me.CodInventario & vbCr & _
                          "    Cantidad límite: " & Format(Abs(CantLimite), fmt) & _
                          "    Cantidad utilizada: " & Format(cantOtras, fmt) & _
                          "    Saldo: " & Format(Saldo, fmt)
                    Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", msg
                End If
            End If
        End If
    End If
    
    'Si la transacción está configurado para verificar límite de cantidad
    If mProps.objGNComprobante.GNTrans.IVVerificaLimite And Not mProps.objGNComprobante.GNTrans.IVControlaSoloCantidad Then
        Dim cantGrabada As Currency
        'Dim Saldo As Currency
        'Dim fmt As String
        cantOtras = 0
        CantLimite = 0
        'Obtiene cantidad del mismo item en otras filas del mismo comprobante
        cantOtras = SumaItemRepetido
        'Verifica el límite de cantidad y cantidad grabada

                If mProps.objGNComprobante.GNTrans.IVControlaLimitePorFactura Then
                    If mProps.objGNComprobante.NumTrans = 0 Then
                        mProps.objGNComprobante.VerificarLimiteItemConTrans mProps.objGNComprobante.CodTrans, mProps.objGNComprobante.NumTrans, mProps.objGNComprobante.IdTransFuente, mProps.CodInventario, CantLimite, cantGrabada
                    Else
                        mProps.objGNComprobante.VerificarLimiteItemConTrans mProps.objGNComprobante.CodTrans, 0, mProps.objGNComprobante.IdTransFuente, mProps.CodInventario, CantLimite, cantGrabada
                    End If
                    If CantLimite = 0 Then
                        Saldo = 0
                    Else
                        If cantGrabada > 0 Then
                            If mProps.objGNComprobante.GNTrans.IVTipoTrans = "I" Then
                                Saldo = cantGrabada - CantLimite
                            Else
                                Saldo = CantLimite + cantGrabada
                            End If
                        ElseIf cantGrabada = 0 Then
                            If mProps.objGNComprobante.GNTrans.IVTipoTrans = "I" Then
                                Saldo = cantGrabada + CantLimite
                            Else
                                Saldo = Abs(CantLimite)
                            End If
                        Else
'                            If mProps.objGNComprobante.GNTrans.IVTipoTrans = "I" Then
'                                Saldo = CantLimite - cantGrabada
'                            Else
                                Saldo = cantGrabada - CantLimite
'                            End If
                        End If

                            'solo  transacciones iguales Ej: Egreso / Egreso
                            '                                Ingreso/ Ingreso
                            If mProps.objGNComprobante.GNTrans.IVTipoTrans = "I" Then
                                'If saldo < 0 Then saldo = 0 modificado jeaa 17/01/05
                                If Saldo < 0 Then
                                    Saldo = Abs(Saldo)
                                End If
                            Else
'                                If Saldo > 0 Then Saldo = 0
                            End If
'                        If mProps.objGNComprobante.GNTrans.IVTipoTrans = "I" Then
                        If Saldo < 0 Then
                            Saldo = (Saldo) + cantOtras     'Devuelve sin signo
                        ElseIf Saldo = 0 Then
                            Saldo = Abs(Saldo) + cantOtras     'Devuelve sin signo
                        Else
                            Saldo = Abs(Saldo) - cantOtras     'Devuelve sin signo
                        End If
'                        End If
                    End If
                    
                    'Si está superando el límite, saca mensaje

                        If Abs(mProps.Cantidad) > Saldo Then
                            fmt = mProps.objGNComprobante.Empresa.GNOpcion.FormatoCantidad
                                msg = "Ha intentado registrar la cantidad mayor al límite " & _
                                      "establecido en la transacción '" & _
                                        mProps.objGNComprobante.GNTrans.IVVerificaLimiteCon & "' y '" & _
                                        mProps.objGNComprobante.GNTrans.IVVerificaLimiteCon & "M'." & vbCr & vbCr & _
                                      "    Código de item: " & mProps.CodInventario & vbCr & _
                                      "    Cantidad límite: " & Format(Abs(CantLimite), fmt) & _
                                      "    Cantidad utilizada: " & Format(Abs(cantGrabada) + Abs(cantOtras), fmt) & _
                                      "    Saldo: " & Format(Saldo, fmt)
                            MsgBox msg, vbInformation
                            If mProps.objGNComprobante.GNTrans.IVVerificaLimiteNoGrabar Then
                                'Genera un error
                                Err.Raise ERR_LIMITEITEM, "IVKardex.VerificaDatos", msg
                            Else
                                'sino no genera error sólo saca mensaje
                                MsgBox msg, vbInformation
                            End If
                        End If
                    'End If
                        
'------------ opcion 2
                ElseIf mProps.objGNComprobante.GNTrans.IVVerificaLimite And mProps.objGNComprobante.GNTrans.IVactualizaTRansFuente Then
                    mProps.objGNComprobante.VerificarLimiteItem Me.CodInventario, CantLimite, cantGrabada

                    'Devuelve saldo de cantidad para que pueda corregir en la pantalla
                    If CantLimite = 0 Then
                        Saldo = 0
                    Else
                            If cantGrabada > 0 Then
                                Saldo = CantLimite + cantGrabada
                            ElseIf cantGrabada <> 0 Then
                                Saldo = cantGrabada
                            Else
                                Saldo = (CantLimite)
                            End If

                            If mProps.objGNComprobante.GNTrans.IVTipoTrans = "I" Then
                                'If saldo < 0 Then saldo = 0 modificado jeaa 17/01/05
                                If Saldo < 0 Then Saldo = Abs(Saldo)
                            Else
                                If Saldo > 0 Then Saldo = 0
                            End If
                        End If
                        Saldo = Abs(Saldo) - cantOtras     'Devuelve sin signo
                    'Si está superando el límite, genera error
                    If Abs(Me.Cantidad) > Saldo Then
                        fmt = mProps.objGNComprobante.Empresa.GNOpcion.FormatoCantidad
                            fmt = mProps.objGNComprobante.Empresa.GNOpcion.FormatoCantidad
                            msg = "Ha intentado registrar la cantidad mayor al límite " & _
                                  "establecido en la transacción '" & _
                                        mProps.objGNComprobante.GNTrans.IVVerificaLimiteCon & "'." & vbCr & vbCr & _
                                  "    Código de item: " & Me.CodInventario & vbCr & _
                                  "    Cantidad límite: " & Format(Abs(CantLimite), fmt) & _
                                  "    Cantidad utilizada: " & Format(Abs(cantGrabada) + cantOtras, fmt) & _
                                  "    Saldo: " & Format(Saldo, fmt)
                            'Si está configurado para que no permita grabar superando el límite
                            If mProps.objGNComprobante.GNTrans.IVVerificaLimiteNoGrabar Then
                                'Genera un error
                                Err.Raise ERR_LIMITEITEM, "IVKardex.VerificaDatos", msg
                            Else
                                'sino no genera error sólo saca mensaje
                                MsgBox msg, vbInformation
                            End If
                        End If
                        
                ElseIf mProps.objGNComprobante.GNTrans.IVControlaLimitePorFacturaxLinea Then
                    If mProps.objGNComprobante.NumTrans = 0 Then
                        mProps.objGNComprobante.VerificarLimiteItemConTransxOrden mProps.objGNComprobante.CodTrans, mProps.objGNComprobante.NumTrans, mProps.objGNComprobante.IdTransFuente, mProps.CodInventario, CInt(IIf(Len(mProps.Nota1) > 0, mProps.Nota1, 0)), CantLimite, cantGrabada
                    Else
                        mProps.objGNComprobante.VerificarLimiteItemConTransxOrden mProps.objGNComprobante.CodTrans, 0, mProps.objGNComprobante.IdTransFuente, mProps.CodInventario, CInt(IIf(Len(mProps.Nota1) > 0, mProps.Nota1, 0)), CantLimite, cantGrabada
                    End If

                    If CantLimite = 0 Then
                        Saldo = 0
                    Else
                        If cantGrabada > 0 Then
                           If mProps.objGNComprobante.GNTrans.IVTipoTrans = "I" Then
                                Saldo = cantGrabada - CantLimite
                            Else
                                Saldo = CantLimite + cantGrabada
                            End If
                        ElseIf cantGrabada = 0 Then
                            If mProps.objGNComprobante.GNTrans.IVTipoTrans = "I" Then
                                Saldo = cantGrabada + CantLimite
                            Else
                                Saldo = Abs(CantLimite)
                            End If
                        Else
                            Saldo = cantGrabada - CantLimite
                        End If
                            'solo  transacciones iguales Ej: Egreso / Egreso
                            '                                Ingreso/ Ingreso
                            If mProps.objGNComprobante.GNTrans.IVTipoTrans = "I" Then
                                If Saldo < 0 Then
                                    Saldo = Abs(Saldo)
                                End If
                            End If
                        If Saldo < 0 Then
                            Saldo = (Saldo) + cantOtras     'Devuelve sin signo
                        ElseIf Saldo = 0 Then
                            Saldo = Abs(Saldo) + cantOtras     'Devuelve sin signo
                        Else
                            Saldo = Abs(Saldo) - cantOtras     'Devuelve sin signo
                        End If
                    End If
                    'Si está superando el límite, saca mensaje
                        If Abs(mProps.Cantidad) > Saldo Then
                            fmt = mProps.objGNComprobante.Empresa.GNOpcion.FormatoCantidad
                                msg = "Ha intentado registrar la cantidad mayor al límite " & _
                                      "establecido en la transacción '" & _
                                        mProps.objGNComprobante.GNTrans.IVVerificaLimiteCon & "' y '" & _
                                        mProps.objGNComprobante.GNTrans.IVVerificaLimiteCon & "M'." & vbCr & vbCr & _
                                      "    Código de item: " & mProps.CodInventario & vbCr & _
                                      "    Cantidad límite: " & Format(Abs(CantLimite), fmt) & _
                                      "    Cantidad utilizada: " & Format(Abs(cantGrabada) + Abs(cantOtras), fmt) & _
                                      "    Saldo: " & Format(Saldo, fmt)
                            MsgBox msg, vbInformation
                            If mProps.objGNComprobante.GNTrans.IVVerificaLimiteNoGrabar Then
                                'Genera un error
                                Err.Raise ERR_LIMITEITEM, "IVKardex.VerificaDatos", msg
                            Else
                                'sino no genera error sólo saca mensaje
                                MsgBox msg, vbInformation
                            End If
                        End If
                        
                    '---------------- final
                    Else
                        'jeaa 22/04/2009
                        'cuando son de varias transacciones fuentes
                        If Len(mProps.objGNComprobante.idCentros) > 0 Then
                            mProps.objGNComprobante.VerificarLimiteItemConVariasTransaccionesFuente Me.CodInventario, CantLimite, cantGrabada
                        Else
                            mProps.objGNComprobante.VerificarLimiteItem Me.CodInventario, CantLimite, cantGrabada
                        End If
    
                        'Devuelve saldo de cantidad para que pueda corregir en la pantalla
                        If CantLimite = 0 Then
                            Saldo = 0
                        Else
                                Saldo = CantLimite + cantGrabada
    
                                If mProps.objGNComprobante.GNTrans.IVTipoTrans = "I" Then
                                    'If saldo < 0 Then saldo = 0 modificado jeaa 17/01/05
                                    If Saldo < 0 Then Saldo = Abs(Saldo)
                                Else
                                    If Saldo > 0 Then Saldo = 0
                                End If
                            End If
                            Saldo = Abs(Saldo) - cantOtras     'Devuelve sin signo
                        'Si está superando el límite, genera error
                        If Abs(Me.Cantidad) > Saldo Then
                            fmt = mProps.objGNComprobante.Empresa.GNOpcion.FormatoCantidad
                                fmt = mProps.objGNComprobante.Empresa.GNOpcion.FormatoCantidad
                                msg = "Ha intentado registrar la cantidad mayor al límite " & _
                                      "establecido en la transacción '" & _
                                            mProps.objGNComprobante.GNTrans.IVVerificaLimiteCon & "'." & vbCr & vbCr & _
                                      "    Código de item: " & Me.CodInventario & vbCr & _
                                      "    Cantidad límite: " & Format(Abs(CantLimite), fmt) & _
                                      "    Cantidad utilizada: " & Format(Abs(cantGrabada) + cantOtras, fmt) & _
                                      "    Saldo: " & Format(Saldo, fmt)
                                'Si está configurado para que no permita grabar superando el límite
                                If mProps.objGNComprobante.GNTrans.IVVerificaLimiteNoGrabar Then
                                    'Genera un error
                                    Err.Raise ERR_LIMITEITEM, "IVKardex.VerificaDatos", msg
                                Else
                                    'sino no genera error sólo saca mensaje
                                    MsgBox msg, vbInformation
                                End If
                            End If
                        
                    End If
                End If
                
    'verifica precios x Grupo cliente
            'AUC 21/12/06 valida que se ingrese la observacion
    If mProps.objGNComprobante.GNTrans.IVPideObservacion Then
         If Len(mProps.Nota) = 0 And item.Tipo = Preparacion Then
            Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", _
                "La Debe Ingresar Observacion/Num. Piezas.."
         End If
    End If
    
        If mProps.objGNComprobante.GNTrans.IVOmitirIVARISE Then 'AUC agregado  para omitir RISE
            Set pc = mProps.objGNComprobante.Empresa.RecuperaPCProvCli(mProps.objGNComprobante.IdProveedorRef)
            If Not pc Is Nothing Then
                If (pc.TipoProvCli = "RISE" Or pc.TipoProvCli = "ARTE") And (mProps.IVA <> 0) Then
                    mProps.IVA = 0
                End If
            End If
        End If
        If mProps.objGNComprobante.GNTrans.CodPantalla = "IVPVTPROD" Then
            If item.Tipo = rubro Then
                For i = 2 To 5
                    If mProps.IdGrupo(i) = 0 Then
                        Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", _
                        "Debe Especificar Valores Correspondientes a la producción "
                    End If
                Next
            End If
        End If
        'AUC ESTO YA NO VA  PORQUE SE CAMBIO DE PANTALLA
'        If mProps.objGNComprobante.GNTrans.IVGeneraNumSerie And mProps.objGNComprobante.GNTrans.IVTipoTrans = "E" Then
'            If item.bandSerie Then
'                If Len(mProps.TiempoEntrega) = 0 Or Abs(mProps.Cantidad) > 1 Then
'                    Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", _
'                    "Debe Ingresar Numero de serie o la cantidad no pude ser diferente de 1"
'                End If
'            End If
'        End If
'
'''        If InStr(1, UCase(mProps.objGNComprobante.Empresa.CodEmpresa), "LACTO") > 0 Then
'''
'''            If mProps.Arancel = 0 Then
'''               Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", _
'''                  "El PH no puede ser 0. (" & mProps.CodInventario & ")"
'''            End If
'''
'''            If mProps.ice = 0 Then
'''               Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", _
'''                  "La Grasa no puede ser 0. (" & mProps.CodInventario & ")"
'''            End If
'''
'''            If mProps.FODIN = 0 Then
'''               Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", _
'''                  "La Proteina no puede ser 0. (" & mProps.CodInventario & ")"
'''            End If
'''
'''            If mProps.Comision = 0 Then
'''               Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", _
'''                  "El Agua no puede ser 0. (" & mProps.CodInventario & ")"
'''            End If
'''
'''            If mProps.NumDias = 0 Then
'''               Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", _
'''                  "El QUIO no puede ser 0. (" & mProps.CodInventario & ")"
'''            End If
'''
'''            If Len(mProps.TiempoEntrega) = 0 Then
'''               Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", _
'''                  "El REDUC no puede ser 0. (" & mProps.CodInventario & ")"
'''            End If
'''
'''            If mProps.DescuentoOriginal = 0 Then
'''               Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", _
'''                  "La Merma no puede ser 0. (" & mProps.CodInventario & ")"
'''            End If
'''
'''
'''        End If
        
        Set item = Nothing
End Sub


'Private Function VerificaConFuente(ByRef CantLimite) As Boolean
'    Dim sql As String, rs As Recordset
'
'    sql = "SELECT  Sum(Cantidad) as Total   FROM IVKardex Where TransID = " & _
'           mProps.objGNComprobante.IdTransFuente & " AND IdInventario = '" & _
'           mProps.IdInventario & "' "
'
'    Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
'
'    If rs.EOF Then
'        VerificaConFuente = False
'    Else
'        If IsNull(rs!total) Then
'            VerificaConFuente = False
'        Else
'            CantLimite = Abs(rs!total)
'            VerificaConFuente = True
'        End If
'
'    End If
'    rs.Close
'    Set rs = Nothing
'End Function

Private Function SumaItemRepetido() As Long
    Dim i As Long, ivk As IVKardex, cantOtras As Currency
    
    For i = 1 To mProps.objGNComprobante.CountIVKardex
        Set ivk = mProps.objGNComprobante.IVKardex(i)
        If (Not (ivk Is Me)) And (ivk.CodInventario = Me.CodInventario) Then
            cantOtras = cantOtras + ivk.Cantidad
        End If
    Next i
    Set ivk = Nothing

    SumaItemRepetido = cantOtras
End Function

Friend Sub Grabar()
    Dim sql As String, rs As Recordset
    Dim item As IVInventario
    Dim pc As PCProvCli
    Dim SecuencialItem As Long
    VerificaDatos
    'Asegura que en IVExist haya un registro correspondiente
    AseguraBodega
    '***Agregado. 02/sep/2003
    'Busca Numero de Precio seleccionado
    If Me.GNComprobante.EsNuevo Then
        If Me.GNComprobante.GNTrans.IVCalcularDescPrecioPre Then
            BuscaNumeroPrecioRef 'auc PARA YOLITA
        Else
           If mProps.idpadre <> 0 Then
                BuscaNumeroPrecioPadre mProps.idpadre
            Else
                BuscaNumeroPrecio
            End If
        End If
    End If
    sql = "SELECT * FROM IVKardex WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !transid = mProps.objGNComprobante.transid
        !IdInventario = mProps.IdInventario
        !IdBodega = mProps.IdBodega
        !Cantidad = mProps.Cantidad
        !CostoTotal = mProps.CostoTotal
        !CostoRealTotal = mProps.CostoRealTotal
        !PrecioTotal = mProps.PrecioTotal
        !PrecioRealTotal = mProps.PrecioRealTotal
        !Descuento = mProps.Descuento
        !IVA = mProps.IVA
        !Orden = mProps.Orden
        !Nota = mProps.Nota
        !Nota1 = mProps.Nota1
        !NumeroPrecio = mProps.NumeroPrecio
        !ValorRecargoItem = mProps.ValorRecargoItem '***Agregado. Angel. 29/jul/2004
        !TiempoEntrega = mProps.TiempoEntrega   ' jeaa 22/09/2005
        !IdICE = mProps.IdICE '***Agregado. jeaa 21/07/2006
        !ValorICEItem = mProps.PorcentajeRecargoIce '***Agregado. jeaa 21/07/2006
        If Me.GNComprobante.GNTrans.CodPantalla = "IVISOFAC" Then
                mProps.bandImprimir = False
        End If
        !bandImprimir = mProps.bandImprimir 'AUC 21/07/06
       !idpadre = mProps.idpadre 'AUC 15/09/06
        !bandVer = mProps.bandVer 'AUC 21/07/06
        !IdPadreSub = mProps.IdPadreSub  'AUC 15/09/06
        !FechaDevol = mProps.FechaDevol   'AUC para demostracion de los andamios
        !NumDias = mProps.NumDias
        !FechaLleva = mProps.FechaLleva    'AUC para demostracion de los andamios
        !DescuentoOriginal = mProps.DescuentoOriginal
        !Arancel = mProps.Arancel
        !ice = mProps.ice
        !FODIN = mProps.FODIN
        !Comision = mProps.Comision
        !CostoReferencial = mProps.CostoReferencial
        !Tamaño = mProps.Tamaño
        !RecargoArancel = mProps.RecargoArancel
        If mProps.NumCentro > 0 Then   'NO ESTA GRABADO
            GrabarCentroCosto
            !idCentroCosto = mProps.idCentroCosto
        End If
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.Id = !Id             'Guarda el Id(Autonumerico)
        .Close
        'AUC 23/03/06 agregado bandera para generar o no el kardex de subitems
        With mProps.objGNComprobante.GNTrans
            Set item = Me.GNComprobante.Empresa.RecuperaIVInventario(Me.IdInventario)
              If item.Tipo = 4 And .IVTipoTrans = "E" And .MostrarSubItemsImp Then
                 GeneraPreparacion item 'AUC 01/15/06  para retroburguer
             End If
            If mProps.objGNComprobante.GNTrans.IVActualizaCostoUCenItems And mProps.objGNComprobante.EsNuevo Or mProps.objGNComprobante.GNTrans.IVActualizaCostoUCenItemsMod Then
                 mProps.objGNComprobante.Empresa.GrabarCostoUltIngreso mProps.IdInventario, mProps.CostoRealTotal / mProps.Cantidad
            End If
            If mProps.objGNComprobante.GNTrans.IVActualizaCostoRefItems And mProps.objGNComprobante.EsNuevo Then
                 mProps.objGNComprobante.Empresa.GrabarCostoReferencial mProps.IdInventario, mProps.CostoRealTotal / mProps.Cantidad
            End If
           If .IVTipoTrans = "A" Or .IVTipoTrans = "E" And mProps.objGNComprobante.GNTrans.CodPantalla = "IV_RESALQ" Then
                 GrabarIVReservacion
            End If
        End With
        If Not item.BandServicio Then
            If mProps.objGNComprobante.GNTrans.CodPantalla = "IVBQDISO" And Not mProps.objGNComprobante.GNTrans.IVMotReclamoISO Then
                GrabarIVInventarioDetalleISO
            End If
            If mProps.objGNComprobante.GNTrans.CodPantalla = "IVBDISO" And mProps.Cantidad < 0 Then
                mProps.objGNComprobante.ActualizaBandVerificaNew mProps.TiempoEntrega, mProps.bandVer
                If mProps.objGNComprobante.GNTrans.IVTransCal Then
                    mProps.objGNComprobante.ActualizaTransIdNew mProps.TiempoEntrega, "TransIDCalifica", mProps.objGNComprobante.transid
                    mProps.objGNComprobante.ActualizaTransIdNew mProps.TiempoEntrega, "IdMotivoCal", mProps.IdICE
                End If
                
                
            End If
            If mProps.objGNComprobante.GNTrans.CodPantalla = "IVISOPRD" Then
                If mProps.objGNComprobante.GNTrans.IVDetalleConsumoISO Then
                    GrabaConsumoDetalleOTR
                Else
                    GrabaConsumoDetalle
                End If
                
                
                    mProps.objGNComprobante.ActualizaDatosProcesoNew mProps.TiempoEntrega, mProps.bandImprimir, mProps.IdInventario
                    If mProps.objGNComprobante.GNTrans.IVTransProd Then
                        mProps.objGNComprobante.ActualizaTransIdNew mProps.TiempoEntrega, "TransIDProceso", mProps.objGNComprobante.transid
                        If Not mProps.BandProceso Then
                            mProps.objGNComprobante.ActualizaTransIdNew mProps.TiempoEntrega, "IdMotivoPro", mProps.IdICE
                            mProps.objGNComprobante.ActualizaTransIdNew mProps.TiempoEntrega, "IdProcesado", 0
                            mProps.objGNComprobante.ActualizaTransIdNew mProps.TiempoEntrega, "BandProceso", mProps.BandProceso
                        Else
                            mProps.objGNComprobante.ActualizaTransIdNew mProps.TiempoEntrega, "IdMotivoPro", 0
                        End If
                        If mProps.Motivo = 3 Then
                            mProps.objGNComprobante.ActualizaTransIdN mProps.TiempoEntrega, "ValorTrabajo", mProps.ValorTrabajo
                        End If
                        
                    Else
                    End If
            End If
            If mProps.TiempoEntrega <> "0" And Len(mProps.TiempoEntrega) > 0 Then
                If Not mProps.objGNComprobante.GNTrans.IVTipoTrans = "I" Then
                     If mProps.objGNComprobante.GNTrans.CodPantalla = "IVISOFAC" Then
                             If mProps.objGNComprobante.NumDias = 0 Then
                                If mProps.objGNComprobante.CodTrans = "FVCAR" Then
                                    
                                
                                Else
                                    mProps.objGNComprobante.ActualizaBandFacturaNew mProps.TiempoEntrega, True
                                    mProps.objGNComprobante.ActualizaTransIdNew mProps.TiempoEntrega, "TransIDFactura", mProps.objGNComprobante.transid
                                End If
                             Else
                                 mProps.objGNComprobante.ActualizaTransIdNew mProps.TiempoEntrega, "TransIDFacturaCarc", mProps.objGNComprobante.transid
                                 If mProps.objGNComprobante.CodTrans = "FVCAR" Or mProps.objGNComprobante.CodTrans = "SCI" Or mProps.objGNComprobante.CodTrans = "FSCI" Then
                                    mProps.objGNComprobante.ActualizaTransIdNew mProps.TiempoEntrega, "bandcarcasausada", 0
                                
                                 End If
                             End If
                     End If
                 End If
            
                If mProps.objGNComprobante.GNTrans.CodPantalla = "IVDVISO" And mProps.objGNComprobante.NumDias = 0 Then
                           'mProps.objGNComprobante.ActualizaBandFacturaNew mProps.TiempoEntrega, False
                            'mProps.objGNComprobante.ActualizaTransIdNew mProps.TiempoEntrega, "TransIDFactura", 0
                ElseIf mProps.objGNComprobante.GNTrans.CodPantalla = "IVDVISO" And mProps.objGNComprobante.NumDias = 1 Then
                            'mProps.objGNComprobante.ActualizaTransIdNew mProps.TiempoEntrega, "TransIDFacturaCarc", 0
                End If
        End If
    Else
        If mProps.objGNComprobante.GNTrans.IVTransProd Then
            mProps.objGNComprobante.ActualizaTransIdNew mProps.TiempoEntrega, "TransIDProceso", mProps.objGNComprobante.transid
            If Not mProps.BandProceso Then
                mProps.objGNComprobante.ActualizaTransIdNew mProps.TiempoEntrega, "IdMotivoPro", mProps.IdICE
                mProps.objGNComprobante.ActualizaTransIdNew mProps.TiempoEntrega, "IdProcesado", 0
                mProps.objGNComprobante.ActualizaTransIdNew mProps.TiempoEntrega, "BandProceso", mProps.BandProceso
            Else
                mProps.objGNComprobante.ActualizaTransIdNew mProps.TiempoEntrega, "IdMotivoPro", 0
                mProps.objGNComprobante.ActualizaTransIdNew mProps.TiempoEntrega, "BandProceso", mProps.BandProceso
            End If
            If mProps.Motivo = 3 Then
                mProps.objGNComprobante.ActualizaTransIdN mProps.TiempoEntrega, "ValorTrabajo", mProps.ValorTrabajo
            End If
            
        Else
        End If
    
    End If
    
    If mProps.objGNComprobante.GNTrans.CodPantalla = "IVPVTPROD" Then
        CreaItemProduccion 'ver si esto va ha servir
    End If
    
    If mProps.bandVer Then
        If mProps.objGNComprobante.GNTrans.CodPantalla = "IVOP" Then
            CreaItemProduccionOP
        End If
    End If
    
    If mProps.PrecioTotal <> 0 Then
        GrabarIVFinanciamientoItem
    End If
'    If mProps.objGNComprobante.GNTrans.IVGeneraNumSerie And mProps.objGNComprobante.GNTrans.IVTipoTrans = "E" Then
'        CambiaEstadoIvNumSerie
'    End If
    If mProps.objGNComprobante.GNTrans.IVActualizaPVIVA Then
        ActualizaPrecioItems
    End If
    If mProps.objGNComprobante.GNTrans.IVGeneraNumSerie Then
        AsignaIdIVKardex
    End If
    If mProps.objGNComprobante.GNTrans.CodTrans = "IVBQDPACA" Then
        GrabarIVKardexPaca
    End If
    End With
    Set item = Nothing
    Set rs = Nothing
    Set pc = Nothing
End Sub

Friend Sub Recuperar(rs As Recordset)
Dim i As Integer
    With rs
        If Not .EOF Then
            mProps.Id = !Id
            If Not IsNull(!IdInventario) Then mProps.IdInventario = !IdInventario
            If Not IsNull(!IdBodega) Then mProps.IdBodega = !IdBodega
            If Not IsNull(!CodInventario) Then mProps.CodInventario = !CodInventario
            If Not IsNull(!CodBodega) Then mProps.CodBodega = !CodBodega
            If Not IsNull(!Cantidad) Then mProps.Cantidad = !Cantidad
            
            If Not IsNull(!CostoTotal) Then mProps.CostoTotal = !CostoTotal
            If Not IsNull(!CostoRealTotal) Then mProps.CostoRealTotal = !CostoRealTotal
            If Not IsNull(!PrecioTotal) Then mProps.PrecioTotal = !PrecioTotal
            If Not IsNull(!PrecioRealTotal) Then mProps.PrecioRealTotal = !PrecioRealTotal
            
            If Not IsNull(!Descuento) Then mProps.Descuento = !Descuento
            If Not IsNull(!IVA) Then mProps.IVA = !IVA
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            If Not IsNull(!Nota) Then mProps.Nota = !Nota
            If Not IsNull(!Nota1) Then mProps.Nota1 = !Nota1
            If Not IsNull(!NumeroPrecio) Then mProps.NumeroPrecio = !NumeroPrecio '***Agregado. 02/sep/2003. Angel
            If Not IsNull(!ValorRecargoItem) Then mProps.ValorRecargoItem = !ValorRecargoItem '***Agregado. 29/jul/2004. Angel
            
            If Not IsNull(!IdCuentaActivo) Then mProps.auxIdCuentaActivo = !IdCuentaActivo
            If Not IsNull(!IdCuentaCosto) Then mProps.auxIdCuentaCosto = !IdCuentaCosto
            If Not IsNull(!IdCuentaVenta) Then mProps.auxIdCuentaVenta = !IdCuentaVenta
            If Not IsNull(!IdCuentaDiferida) Then mProps.auxIdCuentaDiferida = !IdCuentaDiferida
            If Not IsNull(!TiempoEntrega) Then mProps.TiempoEntrega = !TiempoEntrega  ' *** jeaa 22/09/2005
            If Not IsNull(!IdICE) Then mProps.IdICE = !IdICE '***Agregado. jeaa 21/07/2006
            If Not IsNull(!ValorICEItem) Then mProps.PorcentajeRecargoIce = !ValorICEItem
            If Not IsNull(!idpadre) Then mProps.idpadre = !idpadre
            If Not IsNull(!bandImprimir) Then mProps.bandImprimir = !bandImprimir  ' *** AUC 08/06/2006
            If Not IsNull(!IdPadreSub) Then mProps.IdPadreSub = !IdPadreSub     ' *** AUC 16/07/2007
            If Not IsNull(!bandVer) Then mProps.bandVer = !bandVer   ' *** AUC 27/06/2006
            If Not IsNull(!codItemPadre) Then mProps.codItemPadre = !codItemPadre
            If Not IsNull(!CodItemPadreSub) Then mProps.CodItemPadreSub = !CodItemPadreSub
            If Not IsNull(!FechaDevol) Then mProps.FechaDevol = !FechaDevol      ' *** AUC para demostracion de los andamios
            If Not IsNull(!NumDias) Then mProps.NumDias = !NumDias
            If Not IsNull(!FechaLleva) Then mProps.FechaLleva = !FechaLleva
            If Not IsNull(!DescuentoOriginal) Then mProps.DescuentoOriginal = !DescuentoOriginal
            If Not IsNull(!Arancel) Then mProps.Arancel = !Arancel
            If Not IsNull(!FODIN) Then mProps.FODIN = !FODIN
            If Not IsNull(!RecargoArancel) Then mProps.RecargoArancel = !RecargoArancel
            If Not IsNull(!ice) Then mProps.ice = !ice
            If Not IsNull(!Comision) Then mProps.Comision = !Comision
            If Not IsNull(!CostoReferencial) Then mProps.CostoReferencial = !CostoReferencial
            If Not IsNull(!Tamaño) Then mProps.Tamaño = !Tamaño
            If Not IsNull(!IdGrupo1) Then mProps.IdGrupo(1) = !IdGrupo1
            If Not IsNull(!IdGrupo2) Then mProps.IdGrupo(2) = !IdGrupo2
            If Not IsNull(!IdGrupo3) Then mProps.IdGrupo(3) = !IdGrupo3
            If Not IsNull(!IdGrupo4) Then mProps.IdGrupo(4) = !IdGrupo4
            If Not IsNull(!IdGrupo5) Then mProps.IdGrupo(5) = !IdGrupo5
            If Not IsNull(!BandCambiaDescr) Then mProps.auxBandCambiaDescr = !BandCambiaDescr
            If Not IsNull(!bandSerie) Then mProps.bandSerie = !bandSerie
            If Not IsNull(!PrecioIVA1) Then mProps.PV(1) = !PrecioIVA1
            If Not IsNull(!PrecioIVA2) Then mProps.PV(2) = !PrecioIVA2

            
            If Not mProps.objGNComprobante.GNTrans.IVMotReclamoISO Then
                If InStr(1, mProps.objGNComprobante.GNTrans.CodPantalla, "ISO") > 0 Then
                    If Len(mProps.TiempoEntrega) > 0 Then
                        RecuperarIVInventarioDetalleISO mProps.TiempoEntrega
                    End If
                End If
            ElseIf mProps.objGNComprobante.GNTrans.CodPantalla = "IVISOFAC" Then
                If InStr(1, mProps.objGNComprobante.GNTrans.CodPantalla, "ISO") > 0 Then
                    If Len(mProps.TiempoEntrega) > 0 Then
                        RecuperarIVInventarioDetalleISO mProps.TiempoEntrega
                    End If
                End If
            
            End If
          'If mProps.objGNComprobante.GNTrans.IVDetalleConsumoISO Then
          CargaConsumoDetalle
          RecuperarIVFinanciamientoItem mProps.objGNComprobante.transid, Me.Id
        End If
    End With
End Sub

Friend Function Clone() As IVKardex
    Dim nuevo As IVKardex
    
    Set nuevo = New IVKardex
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROP)          'Para el método Clone
    mProps = value
    mProps.Id = 0
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Set GNComprobante(value As GNComprobante)
    Set mProps.objGNComprobante = value
    
End Property

Public Property Get GNComprobante() As GNComprobante
    Set GNComprobante = mProps.objGNComprobante
End Property

'AUC cambiado para el reproceso de recetas 19/10/06
Public Property Get Id() As Long
    Id = mProps.Id
End Property

Public Property Get IdSecuencial() As Long
    IdSecuencial = mProps.IdSecuencial
End Property



'Friend Property Let IdInventario(ByVal value As Long)
'    mProps.IdInventario = value
'    mProps.objGNComprobante.SetModificado
'End Property

'AUC cambiado de friend a public porque necesito en siitools
Public Property Get IdInventario() As Long
    IdInventario = mProps.IdInventario
End Property


Public Property Let CodInventario(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodInventario Then
        If Len(value) > 0 Then
            'Actualiza IdInventario también
            sql = "SELECT IdInventario,IdCuentaActivo,IdCuentaCosto,IdCuentaVenta,idcuentadiferida,ValorRecargo, isnull(idice,0) as idice, isnull(porcentajeice,0) as porcentajeice, bandcambiadescr " & _
                  "FROM IVInventario WHERE CodInventario='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "CTLibroDetalle_CodInventario", MSGERR_NOHAYCODIGO
            Else
                mProps.IdInventario = rs!IdInventario
                mProps.auxIdCuentaActivo = rs!IdCuentaActivo
                mProps.auxIdCuentaCosto = rs!IdCuentaCosto
                mProps.auxIdCuentaVenta = rs!IdCuentaVenta
                mProps.auxIdCuentaDiferida = rs!IdCuentaDiferida
                mProps.ValorRecargoItem = rs!ValorRecargo '***Agregado. 29/jul/2004. Angel
                mProps.PorcentajeRecargoIce = rs!PorcentajeICE '***Agregado. 21/07/2006 jeaa
                mProps.IdICE = rs!IdICE '***Agregado. 21/07/2006 jeaa
                mProps.auxBandCambiaDescr = rs!BandCambiaDescr
            End If
            rs.Close
        Else
            mProps.IdInventario = 0
            mProps.auxIdCuentaActivo = 0
            mProps.auxIdCuentaCosto = 0
            mProps.auxIdCuentaVenta = 0
            mProps.auxIdCuentaDiferida = 0
            mProps.ValorRecargoItem = 0 '***Agregado. 29/jul/2004. Angel
            mProps.PorcentajeRecargoIce = 0 '***Agregado. 21/07/2006 jeaa
            mProps.IdICE = 0 '***Agregado. 21/07/2006 jeaa
        End If
    End If

    Set rs = Nothing
    mProps.CodInventario = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodInventario() As String
    CodInventario = Trim$(mProps.CodInventario)
End Property

Friend Property Let IdBodega(ByVal value As Long)
    mProps.IdBodega = value
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Get IdBodega() As Long
    IdBodega = mProps.IdBodega
End Property

Public Property Let CodBodega(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodBodega Then
        If Len(value) > 0 Then
            'Actualiza IdBodega también
            sql = "SELECT IdBodega FROM IVBodega WHERE CodBodega='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "CTLibroDetalle_CodBodega", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                Me.IdBodega = rs!IdBodega
            End If
            rs.Close
        Else
            Me.IdBodega = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodBodega = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodBodega() As String
    CodBodega = Trim$(mProps.CodBodega)
End Property

Public Property Let Cantidad(ByVal value As Currency)
    mProps.Cantidad = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Cantidad() As Currency
    Cantidad = mProps.Cantidad
End Property

Public Property Get Costo() As Currency
    If mProps.Cantidad Then
        Costo = mProps.CostoTotal / mProps.Cantidad / mProps.objGNComprobante.Cotizacion("")
    End If
End Property

Public Property Let CostoTotal(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    mProps.CostoTotal = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CostoTotal() As Currency
    'Devuelve en Moneda ex.
    CostoTotal = mProps.CostoTotal / mProps.objGNComprobante.Cotizacion("")
End Property

Public Property Get CostoReal() As Currency
    If mProps.Cantidad Then
        CostoReal = mProps.CostoRealTotal / mProps.Cantidad / mProps.objGNComprobante.Cotizacion("")
    End If
End Property


Public Property Let CostoRealTotal(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    
    mProps.CostoRealTotal = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CostoRealTotal() As Currency
    'Devuelve en Moneda ex.
    CostoRealTotal = mProps.CostoRealTotal / mProps.objGNComprobante.Cotizacion("")
End Property

Public Property Get Precio() As Currency
    If mProps.Cantidad Then
        Precio = mProps.PrecioTotal / mProps.Cantidad / mProps.objGNComprobante.Cotizacion("")
    End If
End Property

Public Property Let PrecioTotal(ByVal value As Currency)
 'jeaa 10/03/05
    Dim mRedondeo As Integer, mPosPunto As Integer
    mPosPunto = InStr(1, mProps.objGNComprobante.FormatoMoneda, ".")
    If mPosPunto > 0 Then
        mRedondeo = Len(Mid$(mProps.objGNComprobante.FormatoMoneda, mPosPunto, Len(mProps.objGNComprobante.FormatoMoneda) - mPosPunto))
    End If
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    If (mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("RedondeoPT")) = "1" Then
        mProps.PrecioTotal = RedondeadoComun(value, mRedondeo)
    Else
        mProps.PrecioTotal = value
    End If
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get PrecioTotal() As Currency
    'Devuelve en Moneda ex.
    PrecioTotal = mProps.PrecioTotal / mProps.objGNComprobante.Cotizacion("")
End Property

Public Property Get PrecioReal() As Currency
    If mProps.Cantidad Then
        PrecioReal = mProps.PrecioRealTotal / mProps.Cantidad / mProps.objGNComprobante.Cotizacion("")
    End If
End Property

Public Property Let PrecioRealTotal(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    
    mProps.PrecioRealTotal = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get PrecioRealTotal() As Currency
    'Devuelve en Moneda ex.
    PrecioRealTotal = mProps.PrecioRealTotal / mProps.objGNComprobante.Cotizacion("")
End Property

Public Property Let Descuento(ByVal value As Currency)
    mProps.Descuento = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Descuento() As Currency
    Descuento = mProps.Descuento
End Property

Public Property Let IVA(ByVal value As Currency)
    mProps.IVA = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get IVA() As Currency
    IVA = mProps.IVA
End Property

Public Property Let Orden(ByVal value As Integer)
    mProps.Orden = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property

Public Property Let Nota(ByVal value As String)
    value = RTrim$(value)
    If Len(value) > 200 Then                 '*** MAKOTO 16/oct/00 Modifcado 40-->80
        Err.Raise ERR_INVALIDO, "IVKardex_Nota", MSGERR_INVALIDO
    End If
    mProps.Nota = value
    mProps.objGNComprobante.SetModificado
End Property
Private Sub Class_Initialize()
    With mProps
        .FechaEntrada = Date
        .FechaAdicional1 = Date
        .FechaAdicional2 = Date
        .FechaAdicional3 = Date
        .auxBandCambiaDescr = False
    End With
End Sub

Public Property Get Nota() As String
    Nota = RTrim$(mProps.Nota)
End Property

Public Property Let TiempoEntrega(ByVal value As String)
    value = RTrim$(value)
    If Len(value) > 40 Then                 '*** MAKOTO 16/oct/00 Modifcado 40-->80
        Err.Raise ERR_INVALIDO, "IVKardex_TiempoEntrega", MSGERR_INVALIDO
    End If
    
    mProps.TiempoEntrega = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get TiempoEntrega() As String
    
    TiempoEntrega = RTrim$(mProps.TiempoEntrega)
End Property


Public Property Get auxIdCuentaActivo() As Long
    'Si no está asignado IdInventario no devuelve nada
    If IdInventario = 0 Then Exit Property
    auxIdCuentaActivo = mProps.auxIdCuentaActivo
End Property

Public Property Get auxIdCuentaCosto() As Long
    'Si no está asignado CodBanco no devuelve nada
    If IdInventario = 0 Then Exit Property
    auxIdCuentaCosto = mProps.auxIdCuentaCosto
End Property

Public Property Get auxIdCuentaVenta() As Long
    'Si no está asignado IdInventario no devuelve nada
    If IdInventario = 0 Then Exit Property
    auxIdCuentaVenta = mProps.auxIdCuentaVenta
End Property

'Friend Sub SetIdInventarioAux(auxIdCuentaActivo As Long, _
'                               auxIdCuentaCosto As Long, _
'                               auxIdCuentaVenta As Long)
'    mProps.auxIdCuentaActivo = auxIdCuentaActivo
'    mProps.auxIdCuentaCosto = auxIdCuentaCosto
'    mProps.auxIdCuentaVenta = auxIdCuentaVenta
'End Sub

Private Sub Class_Terminate()
    Set mProps.objGNComprobante = Nothing
End Sub

'*** MAKOTO 13/dic/00
Public Function CalcularIvaItem( _
                    ByRef por As Double, _
                    ByRef bandCalculado As Boolean) As Currency
    Dim i As Long, RecAfectaIvaItem As Currency
    Dim neto As Currency, v As Currency
    If Not bandCalculado Then   'Para evitar que se repita muchas veces el mismo cálculo cuando llama del control IVRec
        'Sumar(restar) los recargos/descuentos que afecta a IVA por item
        RecAfectaIvaItem = Me.GNComprobante.IVKardexRecAfectaIvaItem
        por = 0
        If Me.GNComprobante.GNTrans.CodPantalla = "IVRES_ALQ" Or Me.GNComprobante.GNTrans.CodPantalla = "IVBQD2BE" Then
            neto = Abs(Me.GNComprobante.IVKardexTotal1(False))     'Total NETO sin recargo prorateado
            neto = neto - Me.GNComprobante.IVKardexDescItemTotal2Bob
        Else
            neto = Abs(Me.GNComprobante.IVKardexTotal(False))    'Total NETO sin recargo prorateado
            neto = neto - Me.GNComprobante.IVKardexDescItemTotal
        End If


        If neto <> 0 Then por = RecAfectaIvaItem / neto
        bandCalculado = True     'Para saltar éste cálculo desde siguiente vez
    End If
    'Decide si calcula en base de Precio o Costo
    If Me.GNComprobante.GNTrans.IVRecargoEnCosto Then
'        v = Abs(Me.CostoTotal) * (1# - Me.Descuento) * Me.IVA
        v = Me.CostoTotal * (1# - Me.Descuento) * Me.IVA        '*** MAKOTO 26/ene/01 Mod.
    Else
'        v = Abs(Me.PrecioRealTotal) * Me.IVA
'        v = Abs(Me.PrecioTotal) * (1# - Me.Descuento) * Me.IVA
        v = Me.PrecioTotal * (1# - Me.Descuento) * Me.IVA       '*** MAKOTO 26/ene/01 Mod.
    End If
    '(v * por) es la parte de Recargo/Desc que afecta a IVA de item
    v = v + (v * por)
    'Redondea al formato de visualización       '*** MAKOTO 01/mar/01 Agregado
    If Me.GNComprobante.GNTrans.Empresa.GNOpcion.ObtenerValor("CalculoIVA4Decimales") = "1" Then
        v = MiCCur(Format$(v, "#0.0000")) 'jeaa 09/03/2007 siempre con 4 decimales
    Else
        v = MiCCur(Format$(v, Me.GNComprobante.FormatoMoneda))
    End If
    CalcularIvaItem = v
End Function

'***Agregado. 02/sep/2003. Angel
Public Property Let NumeroPrecio(value As Byte)
    If (value < 0) Or (value > 7) Then value = 0
    mProps.NumeroPrecio = value
End Property

'***Agregado. 02/sep/2003. Angel
Public Property Get NumeroPrecio() As Byte
    NumeroPrecio = mProps.NumeroPrecio
End Property

'***Agregado. 02/sep/2003. Angel
'***Para saber con que precio fue grabada la transacción
Private Sub BuscaNumeroPrecio()
    Dim iv As IVInventario
    
    Set iv = Me.GNComprobante.Empresa.RecuperaIVInventario(Me.IdInventario)
    If Not (iv Is Nothing) Then
        Select Case Me.Precio
        Case 0
            Me.NumeroPrecio = 0
        Case iv.Precio(1)
            Me.NumeroPrecio = 1
        Case iv.Precio(2)
            Me.NumeroPrecio = 2
        Case iv.Precio(3)
            Me.NumeroPrecio = 3
        Case iv.Precio(4)
            Me.NumeroPrecio = 4
        Case iv.Precio(5)
            Me.NumeroPrecio = 5
        Case iv.Precio(6)
            Me.NumeroPrecio = 6
        Case iv.Precio(7)
            Me.NumeroPrecio = 7
        
        Case Else
            Me.NumeroPrecio = 0
        End Select
    End If
    Set iv = Nothing
End Sub

'***Agregado. 29/jul/2004. Angel
Public Property Let ValorRecargoItem(ByVal value As Currency)
    mProps.ValorRecargoItem = value
    mProps.objGNComprobante.SetModificado
End Property

'***Agregado. 29/jul/2004. Angel
Public Property Get ValorRecargoItem() As Currency
    ValorRecargoItem = mProps.ValorRecargoItem
End Property

'***Agregado. 25/07/2006 jeaa
Public Property Let IdICE(ByVal value As Long)
    mProps.IdICE = value
    mProps.objGNComprobante.SetModificado
End Property

'***Agregado. 25/07/2006 jeaa
Public Property Get IdICE() As Long
    IdICE = mProps.IdICE
End Property



'jeaa 22/03/2005
Public Function RedondeadoComun(ByVal Valor As Currency, ByVal NumDecimal As Integer) As Currency
    Dim ParteEntera As Double, ParteDecimal As Double
    Dim NumRedondear As Integer, num As Integer, ajuste As Double
    Dim numFinal As Double
    Dim snum As String, svalor As String
    Dim bandNegativo As Boolean
    bandNegativo = False
    If Valor < 0 Then
        bandNegativo = True
    End If
    ParteEntera = Int(Abs(Valor))
    ParteDecimal = Abs(Valor) - ParteEntera
    'convierte en texto yy elina entero y punto 0.123 -> '123'
    svalor = Mid$(Trim$(Str(ParteDecimal)), 2, Len(Trim$(Str(ParteDecimal)))) & "000000"
    snum = Mid$(svalor, NumDecimal + 1, 1)
    num = CInt(snum)
    If num >= 5 Then
                        snum = Mid$(svalor, NumDecimal + 1, 1)
                        NumRedondear = CInt(snum)
                        ajuste = 10 - NumRedondear
                        numFinal = ParteEntera + ParteDecimal + (ajuste / (1 * 10 ^ (NumDecimal + 1)))
    Else
                        snum = Mid$(svalor, NumDecimal + 1, 1)
                        NumRedondear = CInt(snum)
                        numFinal = ParteEntera + ParteDecimal - (NumRedondear / (1 * 10 ^ (NumDecimal + 1)))
    End If
    If bandNegativo Then
        numFinal = numFinal * -1
    End If
    RedondeadoComun = numFinal
End Function

Public Function CalculaNumDecimales() As Integer
 'jeaa 10/03/05
    Dim mRedondeo As Integer, mPosPunto As Integer
    mPosPunto = InStr(1, mProps.objGNComprobante.FormatoMoneda, ".")
    If mPosPunto > 0 Then
        mRedondeo = Len(Mid$(mProps.objGNComprobante.FormatoMoneda, mPosPunto, Len(mProps.objGNComprobante.FormatoMoneda) - mPosPunto))
    End If
    CalculaNumDecimales = mRedondeo
End Function

Friend Sub GrabarPreparacion(ByRef j As Long)
    Dim sql As String, rs As Recordset
    'VerificaDatos
    'Asegura que en IVExist haya un registro correspondiente
    'AseguraBodega
    '***Agregado. 02/sep/2003
    'Busca Numero de Precio seleccionado
    'If Me.GNComprobante.EsNuevo Then BuscaNumeroPrecio
    sql = "SELECT * FROM IVKardex WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !transid = mProps.objGNComprobante.transid
        !IdInventario = mProps.objGNComprobante.IVKardex(j).IdInventario
        !IdBodega = mProps.objGNComprobante.IVKardex(j).IdBodega
        !Cantidad = mProps.objGNComprobante.IVKardex(j).Cantidad
        !CostoTotal = mProps.objGNComprobante.IVKardex(j).CostoTotal
        !CostoRealTotal = mProps.objGNComprobante.IVKardex(j).CostoRealTotal
        !PrecioTotal = mProps.objGNComprobante.IVKardex(j).PrecioTotal
        !PrecioRealTotal = mProps.objGNComprobante.IVKardex(j).PrecioRealTotal
        !Descuento = mProps.objGNComprobante.IVKardex(j).Descuento
        !IVA = mProps.objGNComprobante.IVKardex(j).IVA
        !Orden = mProps.objGNComprobante.IVKardex(j).Orden
        !Nota = mProps.objGNComprobante.IVKardex(j).Nota
        !NumeroPrecio = mProps.objGNComprobante.IVKardex(j).NumeroPrecio
        !ValorRecargoItem = mProps.objGNComprobante.IVKardex(j).ValorRecargoItem
        !TiempoEntrega = mProps.objGNComprobante.IVKardex(j).TiempoEntrega
        !IdICE = mProps.objGNComprobante.IVKardex(j).IdICE
        !Arancel = mProps.objGNComprobante.IVKardex(j).Arancel
        !FODIN = mProps.objGNComprobante.IVKardex(j).FODIN
        !ice = mProps.objGNComprobante.IVKardex(j).ice
        !Comision = mProps.objGNComprobante.IVKardex(j).Comision
        !CostoReferencial = mProps.objGNComprobante.IVKardex(j).CostoReferencial
        !ValorFinanciamiento = mProps.objGNComprobante.IVKardex(j).ValorFinanciamiento
        !RecargoArancel = mProps.objGNComprobante.IVKardex(j).RecargoArancel
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
         mProps.Id = !Id             'Guarda el Id(Autonumerico)
        .Close
    End With
    Set rs = Nothing
End Sub
'AUC agregado para poder pasar cuando grabo la receta
Public Property Let IdInventario(ByVal value As Long)
    mProps.IdInventario = value
    mProps.objGNComprobante.SetModificado
End Property

'AUC 05/01/06 agregado para saber si es receta de retroburguer baje los items de relacion
Private Sub GeneraPreparacion(item As IVInventario)
    Dim i As Long, j As Long, a As String, itemEgreso   As IVInventario
    Dim item1 As IVInventario
    Dim CantidadReceta As Long, indice As Long
    Dim ivk As IVKardex, ivktemp As Collection, c As Currency
    Dim CostoTotal As Currency
'    For i = 1 To Me.CountIVKardex
'         Set item = Me.Empresa.RecuperaIVInventario(Me.IVKardex(i).CodInventario)
           'If Me.IVKardex(i).CodInventario <> "" Then
            If item.NumFamiliaDetalle > 0 Then
                If Not (item Is Nothing) Then
                    CantidadReceta = Me.Cantidad
                    If item.Tipo = Preparacion Then    'Solo para Preparacion
                        For j = 1 To item.NumFamiliaDetalle 'ciclo para detalles de items
                            indice = Me.GNComprobante.AddIVKardex
                            Me.GNComprobante.IVKardex(indice).CodInventario = item.RecuperaDetalleFamilia(j).CodInventario
                            Me.GNComprobante.IVKardex(indice).Cantidad = (item.RecuperaDetalleFamilia(j).Cantidad * CantidadReceta)
                            Me.GNComprobante.IVKardex(indice).IdInventario = item.RecuperaDetalleFamilia(j).IdInventario
                            c = item.CostoDouble2Prep(Me.GNComprobante.FechaTrans, 0, 0, Me.GNComprobante.HoraTrans, Me.GNComprobante.IVKardex(indice).IdInventario)
                            'Si el costo calculado está en otra moneda, convierte en moneda de trans.
                            If Me.GNComprobante.CodMoneda <> item.CodMoneda Then
                                c = c * Me.GNComprobante.Cotizacion(item.CodMoneda) / Me.GNComprobante.Cotizacion("")
                            End If
                            Me.GNComprobante.IVKardex(indice).CostoRealTotal = Me.GNComprobante.IVKardex(indice).Cantidad * c
                            Dim sql As String
                            'AUC Aqui se llama a la funcion que va a restar existencia en items de preparacion
                           ' RestaExistenciaPreparacion "-", Me.IVKardex(indice).IdInventario, Abs(Me.IVKardex(indice).Cantidad)
                            Me.GNComprobante.IVKardex(indice).GrabarPreparacion (indice)
                            'desde aqui recursivo
                            Set item1 = Me.GNComprobante.Empresa.RecuperaIVInventario(Me.GNComprobante.IVKardex(indice).CodInventario)
                            'ExisteParaTransformarPrepara item1, indice
                          '  Me.GNComprobante.RemoveIVKardex (indice)
                            'Me.GNComprobante.Ver
'                            If ExisteParaTransformarPrepara(item1, indice) < 0 Then
                               If item1.Tipo = Preparacion Then
                               'MsgBox item1.Descripcion
                                 Me.GNComprobante.RemoveIVKardex (indice)
                                    GeneraPreparacion item1
                                    Set item1 = Nothing
                                End If
'                            End If
                            'Me.GNComprobante.RemoveIVKardex (indice)
                        Next j
                    End If
                End If
            Else
                Exit Sub
            End If
         'End If
    'Next i
'    Set Item = Nothing
End Sub

''AUC 29/12/05 Verifica existencia cuando se hace la transformacion en preparacion
Public Sub ExisteParaTransformarPrepara1(item As IVInventario)
    Dim i As Long, ivk As IVKardex, j As Long, k As Long, ivk2 As IVKardex
    Dim exist As Currency, sumaCant As Currency, indice As Long, IVKpre As IVKardex
    Dim canTotal As Currency
    Dim cant_ori As Currency
    Dim idItem As Long
    Dim ItemPrepara As IVInventario
    If Not (item Is Nothing) Then
        If item.Tipo <> Preparacion Then        'Solo para Preparacion
            Me.GNComprobante.IVKardex(indice).IdInventario = item.RecuperaDetalleFamilia(k).IdInventario
            Set ItemPrepara = Me.GNComprobante.Empresa.RecuperaIVInventario(Me.GNComprobante.IVKardex(indice).IdInventario)
            If Not (ItemPrepara Is Nothing) Then
                If Not ItemPrepara.BandServicio Then
                    exist = ItemPrepara.Existencia(Me.GNComprobante.IVKardex(indice).CodBodega)
                    sumaCant = 0
                    'Obtiene la cantidad original en caso de modificación
                    cant_ori = Me.GNComprobante.VerificarExistencia_CantOrig(Me.Id)
                End If
                Set ivk2 = Nothing
                'Si se hace negativa al sumar
                canTotal = exist + (Me.GNComprobante.IVKardex(indice).Cantidad * Me.GNComprobante.IVKardex(indice).Cantidad) + cant_ori
                If canTotal < 0 Then
                Err.Raise ERR_INVALIDO, "GNComprobante.ExisteParaTransformar", _
                "No tiene Items para generar esta transaccion." & vbCr & vbCr & _
                "Item: " & ItemPrepara.CodInventario & " " & ItemPrepara.Descripcion & vbCr & _
                "Existencia: " & exist & vbCr & _
                "Cantidad: " & Me.GNComprobante.IVKardex(indice).Cantidad & vbCr & _
                "Existencia nueva: " & canTotal & vbCr & _
                "Verifique la existencia .............."
            End If
        End If
            Me.GNComprobante.RemoveIVKardex (indice)
        End If
    End If
    Set ivk = Nothing
    Set ivk2 = Nothing
    Set IVKpre = Nothing
    Set ItemPrepara = Nothing
End Sub

Private Function ExisteParaTransformarPrepara(ByRef item As IVInventario, ByRef indice As Long) As Currency
    Dim i As Long, ivk As IVKardex, j As Long, ivk2 As IVKardex
    Dim exist As Currency, sumaCant As Currency
    Dim cant_ori As Currency
    Dim Valor As Currency
        exist = item.ExistenciaPreparacion(item.IdInventario)
        If Not (item Is Nothing) Then
            If Not item.BandServicio And item.Tipo <> Preparacion Then
                sumaCant = 0
                'Obtiene la cantidad original en caso de modificación
                cant_ori = Me.GNComprobante.VerificarExistencia_CantOrig(Me.Id)
                'Si se hace negativa al sumar
                Valor = exist + Me.Cantidad '+ cant_ori
                If Valor < 0 Then
                    Me.GNComprobante.RemoveIVKardex (indice)
                    Err.Raise ERR_INVALIDO, "GNComprobante.VerificarDatos", _
                        "Al grabar la transacción la existencia se hace negativa, por lo que no puede grabarla." & vbCr & vbCr & _
                        "Item: " & item.CodInventario & " " & item.Descripcion & vbCr & _
                        "Existencia: " & exist & vbCr & _
                        "Cantidad: " & Me.Cantidad & vbCr & _
                        "Existencia nueva: " & exist + Me.Cantidad
                End If
            End If
        End If
    Set ivk = Nothing
    ExisteParaTransformarPrepara = exist
End Function

Public Property Let CodICE(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodICE Then
        If Len(value) > 0 Then
            'Actualiza IdInventario también
            If InStr(1, UCase(mProps.objGNComprobante.Empresa.CodEmpresa), "ISO") = 0 Then
                sql = "SELECT IdICE, CodICE, isnull(Porcentaje,0) as PorcentajeICE " & _
                      "FROM IVRecargoICE WHERE CodICE='" & value & "'"
                Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
                If rs.EOF Then
                    Err.Raise ERR_NOHAYCODIGO, "IvRecargoICE_CodICE", MSGERR_NOHAYCODIGO
                Else
                    mProps.IdICE = rs!IdICE
                    mProps.CodICE = rs!CodICE '***Agregado. 21/07/2006 jeaa
                End If
                rs.Close
            Else
                sql = "SELECT IdMotivo, codmotivo  " & _
                      "FROM Motivo WHERE CodMotivo='" & value & "'"
                Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
                If rs.EOF Then
                    Err.Raise ERR_NOHAYCODIGO, "IvRecargoICE_CodICE", MSGERR_NOHAYCODIGO
                Else
                    mProps.IdICE = rs!IdMotivo
                    mProps.CodICE = rs!CodMotivo
                End If
                rs.Close
            
            End If
        Else
            mProps.IdICE = 0 '***Agregado. 21/07/2006 jeaa
            mProps.CodICE = 0 '***Agregado. 21/07/2006 jeaa
        End If
    End If

    Set rs = Nothing
    mProps.CodICE = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodICE() As String
    CodICE = Trim$(mProps.CodICE)
End Property

Public Property Let PorcentajeRecargoIce(ByVal value As Currency)
    mProps.PorcentajeRecargoIce = value
    mProps.objGNComprobante.SetModificado
End Property

'***Agregado. 29/jul/2004. Angel
Public Property Get PorcentajeRecargoIce() As Currency
    PorcentajeRecargoIce = mProps.PorcentajeRecargoIce
End Property

'*** jeaa 25/09/2006
Public Function CalcularSumaIvaItem( _
                    ByRef por As Double, _
                    ByRef bandCalculado As Boolean) As Currency
    Dim i As Long, RecAfectaIvaItem As Currency
    Dim neto As Currency, v As Currency
    
    If Not bandCalculado Then   'Para evitar que se repita muchas veces el mismo cálculo cuando llama del control IVRec
        'Sumar(restar) los recargos/descuentos que afecta a IVA por item
        RecAfectaIvaItem = Me.GNComprobante.IVKardexRecAfectaIvaItem
        
        por = 0
        neto = Abs(Me.GNComprobante.IVKardexTotal(False))    'Total NETO sin recargo prorateado
        neto = neto - Me.GNComprobante.IVKardexDescItemTotal
        If neto <> 0 Then por = RecAfectaIvaItem / neto
        
        bandCalculado = True     'Para saltar éste cálculo desde siguiente vez
    End If
    
    'Decide si calcula en base de Precio o Costo
    If Me.GNComprobante.GNTrans.IVRecargoEnCosto Then
'        v = Abs(Me.CostoTotal) * (1# - Me.Descuento) * Me.IVA
        If Me.IVA <> 0 Then
            v = Me.CostoTotal * (1# - Me.Descuento)        '*** MAKOTO 26/ene/01 Mod.
        End If
    Else
'        v = Abs(Me.PrecioRealTotal) * Me.IVA
'        v = Abs(Me.PrecioTotal) * (1# - Me.Descuento) * Me.IVA
        If Me.IVA <> 0 Then
            v = Me.PrecioTotal * (1# - Me.Descuento)        '***
        End If
    End If
    
    '(v * por) es la parte de Recargo/Desc que afecta a IVA de item
    v = v + (v * por)
    
    'Redondea al formato de visualización       '*** MAKOTO 01/mar/01 Agregado
    If Me.GNComprobante.GNTrans.Empresa.GNOpcion.ObtenerValor("CalculoIVA4Decimales") = "1" Then
        v = MiCCur(Format$(v, "#0.0000"))
        'v = Round(v + (v * por), 4)
    Else
        v = MiCCur(Format$(v, Me.GNComprobante.FormatoMoneda))
    End If
    
    CalcularSumaIvaItem = v
End Function

'AUC obtiene el valor des descuento
Public Property Get ValorDescuento() As Currency
    ValorDescuento = Abs(mProps.PrecioTotal - mProps.PrecioRealTotal)
End Property

'AUC 27/09/07
Public Property Let idpadre(ByVal value As Long)
    mProps.idpadre = value
End Property

Public Property Get idpadre() As Long
    idpadre = mProps.idpadre
End Property

'jeaa 18/12/2007
Public Property Get ValorCDescuento() As Currency
    ValorCDescuento = Abs(mProps.CostoTotal - mProps.CostoRealTotal)
End Property

Public Property Get CostoU(cantPadre As Currency) As Currency
    If mProps.Cantidad Then
        CostoU = mProps.CostoTotal / cantPadre / mProps.objGNComprobante.Cotizacion("")
    End If
End Property

'Auc Grabar items reservados
Private Sub GrabarIVReservacion()
Dim sql As String
Dim rs As Recordset
sql = "SELECT * FROM ivreservacion WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !transid = mProps.objGNComprobante.transid
        !IdInventario = mProps.IdInventario
        !IdBodega = mProps.IdBodega
        !cant = Abs(mProps.Cantidad)
        !Fechaalquilar = mProps.objGNComprobante.FechaAlquiler
        !FechaDevolucion = mProps.objGNComprobante.FechaDevol
        If mProps.objGNComprobante.GNTrans.IVTipoTrans = "A" Then
            !Estado = ESTADO_RESERVADO
        End If
        .Update
    End With
    Set rs = Nothing
End Sub

'AUC 08/06/06
Public Property Let bandImprimir(value As Boolean)
    mProps.bandImprimir = value
End Property

'***Agregado. 02/sep/2003. Angel
Public Property Get bandImprimir() As Boolean
    bandImprimir = mProps.bandImprimir
End Property

'AUC 27/09/06
'Public Property Let idpadre(ByVal value As Long)
'    mProps.idpadre = value
''    mProps.objGNComprobante.SetModificado
'End Property

'Public Property Get idpadre() As Long
'    idpadre = mProps.idpadre
'End Property

'AUC 27/10/06
Public Property Let bandVer(value As Boolean)
    mProps.bandVer = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get bandVer() As Boolean
    bandVer = mProps.bandVer
End Property

'AUC 16/08/07
Public Property Let IdPadreSub(ByVal value As Long)
    mProps.IdPadreSub = value
End Property
Public Property Get IdPadreSub() As Long
    IdPadreSub = mProps.IdPadreSub
End Property

Public Property Let codItemPadre(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.codItemPadre Then
        If Len(value) > 0 Then
            'Actualiza IdInventario también
            sql = "SELECT idInventario, CodInventario " & _
                  "FROM IVInventario WHERE CodInventario='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IvKardex_CodItemPadre", MSGERR_NOHAYCODIGO
            Else
                mProps.idpadre = rs!IdInventario
                mProps.codItemPadre = rs!CodInventario
            End If
            rs.Close
        Else
            mProps.idpadre = 0
            mProps.codItemPadre = 0
        End If
    End If

    Set rs = Nothing
    mProps.codItemPadre = value
    mProps.objGNComprobante.SetModificado
End Property


Public Property Get codItemPadre() As String
    codItemPadre = Trim$(mProps.codItemPadre)
End Property

'***Agregado. 31/01/2009
'***Para saber con que precio fue grabada la transacción
Private Sub BuscaNumeroPrecioPadre(idP As Long)
    Dim i As Long
    For i = 1 To Me.GNComprobante.CountIVKardex
        If Me.GNComprobante.IVKardex(i).IdInventario = idP Then
            Me.NumeroPrecio = Me.GNComprobante.IVKardex(i).NumeroPrecio
            Exit For
        End If
    Next i
End Sub

Public Property Let CodItemPadreSub(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodItemPadreSub Then
        If Len(value) > 0 Then
            'Actualiza IdInventario también
            sql = "SELECT idInventario, CodInventario " & _
                  "FROM IVInventario WHERE CodInventario='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IvKardex_CodItemPadreSub", MSGERR_NOHAYCODIGO
            Else
                mProps.IdPadreSub = rs!IdInventario
                mProps.CodItemPadreSub = rs!CodInventario
            End If
            rs.Close
        Else
            mProps.IdPadreSub = 0
            mProps.CodItemPadreSub = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodItemPadreSub = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodItemPadreSub() As String
    CodItemPadreSub = Trim$(mProps.CodItemPadreSub)
End Property

Public Property Let IDKardex(ByVal value As Long)
    mProps.IDKardex = value
End Property

Public Property Get IDKardex() As Long
    IDKardex = mProps.IDKardex
End Property


Public Property Let IdTamanio(ByVal value As Long)
    mProps.IdTamanio = value
End Property

Public Property Get IdTamanio() As Long
    IdTamanio = mProps.IdTamanio
End Property

Public Property Let CodTamanio(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodTamanio Then
        If Len(value) > 0 Then
            'Actualiza IdInventario también
            sql = "SELECT idInventario, CodInventario " & _
                  "FROM IVInventario WHERE CodInventario='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IvKardex_CodTamanio", MSGERR_NOHAYCODIGO
            Else
                mProps.IdPadreSub = rs!IdInventario
                mProps.CodTamanio = rs!CodInventario
                mProps.IdTamanio = rs!IdInventario
            End If
            rs.Close
        Else
            mProps.IdPadreSub = 0
            mProps.CodTamanio = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodTamanio = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodTamanio() As String
    CodTamanio = Trim$(mProps.CodTamanio)
End Property

Public Property Let IdTrabajo(ByVal value As Long)
    mProps.IdTrabajo = value
End Property

Public Property Get IdTrabajo() As Long
    IdTrabajo = mProps.IdTrabajo
End Property

Public Property Let Secuencial(ByVal value As Long)
    mProps.Secuencial = value
End Property

Public Property Get Secuencial() As Long
    Secuencial = mProps.Secuencial
End Property

Public Property Let CodTrabajo(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodTrabajo Then
        If Len(value) > 0 Then
            'Actualiza IdInventario también
            sql = "SELECT idInventario, CodInventario " & _
                  "FROM IVInventario WHERE CodInventario='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IvKardex_CodTrabajo", MSGERR_NOHAYCODIGO
            Else
                mProps.IdPadreSub = rs!IdInventario
                mProps.CodTrabajo = rs!CodInventario
                mProps.IdTrabajo = rs!IdInventario
            End If
            rs.Close
        Else
            mProps.IdPadreSub = 0
            mProps.CodTrabajo = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodTrabajo = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodTrabajo() As String
    CodTrabajo = Trim$(mProps.CodTrabajo)
End Property

Public Property Let Original(ByVal value As String)
    value = RTrim$(value)
    If Len(value) > 1 Then
        Err.Raise ERR_INVALIDO, "IVKardex_Original", MSGERR_INVALIDO
    End If
    
    If value <> "R" And value <> "O" Then
        Err.Raise ERR_INVALIDO, "IVKardex_Original", MSGERR_INVALIDO
    End If
    
    mProps.Original = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Original() As String
    Original = RTrim$(mProps.Original)
End Property

Public Property Let Radial(ByVal value As String)
    value = RTrim$(value)
    If Len(value) > 1 Then                 '*** MAKOTO 16/oct/00 Modifcado 40-->80
        Err.Raise ERR_INVALIDO, "IVKardex_Radial", MSGERR_INVALIDO
    End If
    If value <> "R" And value <> "C" Then
        Err.Raise ERR_INVALIDO, "IVKardex_Radial", MSGERR_INVALIDO
    End If
    
    mProps.Radial = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Radial() As String
    Radial = RTrim$(mProps.Radial)
End Property

Public Property Let NumSerie(ByVal value As String)
    value = RTrim$(value)
    If Len(value) > 20 Then                 '*** MAKOTO 16/oct/00 Modifcado 40-->80
        Err.Raise ERR_INVALIDO, "IVKardex_NumSerie", MSGERR_INVALIDO
    End If
    
    mProps.NumSerie = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get NumSerie() As String
    NumSerie = RTrim$(mProps.NumSerie)
End Property

Public Property Let DisenioOri(ByVal value As String)
    value = RTrim$(value)
    If Len(value) > 40 Then                 '*** MAKOTO 16/oct/00 Modifcado 40-->80
        Err.Raise ERR_INVALIDO, "IVKardex_DisenioOri", MSGERR_INVALIDO
    End If
    
    mProps.DisenioOri = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get DisenioOri() As String
    DisenioOri = RTrim$(mProps.DisenioOri)
End Property

Private Sub GrabarIVInventarioDetalleISO()
    Dim sql As String, rs As Recordset
    Dim AuxDesc As Integer, SecuencialItem As Long
    Dim rsA As Recordset
   
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM IVInventarioDetalleISO WHERE ID=" & mProps.IdSecuencial
        Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
        If Not rs.EOF Then
            #If DAOLIB Then
                rs.Edit
            #End If
        Else
'            If mProps.Cantidad > 0 Then Exit Sub
            sql = "SELECT * FROM IVInventarioDetalleISO WHERE 1=0"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
            rs.AddNew
        End If
        
    With rs
        !IDKardex = mProps.Id                       'Es el mismo asignado a GnComprobante
        !IdTamanio = mProps.IdInventario
        !Original = mProps.Original
        !Radial = mProps.Radial
        !DisenioOri = mProps.DisenioOri
        !IdTrabajo = mProps.IdTrabajo
        !Secuencial = mProps.Secuencial
        !NumSerie = mProps.NumSerie
        !transid = mProps.objGNComprobante.transid
        If mProps.objGNComprobante.EsNuevo Then
            !BandVerifica = mProps.BandVerifica
            !BandProceso = mProps.BandProceso
            !BandFactura = mProps.BandFactura
            !BandCarcasaISO = mProps.BandCarcasaISO
            !BandCarcasaUsada = mProps.BandCarcasaUsada
            !transid = mProps.objGNComprobante.transid
            !TransIDCalifica = mProps.objGNComprobante.IdTransFuente
            !TransIDProceso = mProps.TransIDProceso
            !TransIDFactura = mProps.TransIDFactura
            !TransIDFacturaCarc = mProps.TransIDFacturaCarc
            !BandComprada = mProps.BandComprada
            !TransIDCompra = mProps.TransIDCompra
            !ValorCarcasa = mProps.ValorCarcasa
            !ValorTrabajo = mProps.ValorTrabajo
            !BandPagado = mProps.BandPagado
        End If
        !IdMotivoCal = mProps.IdMotivoCal
        !IdMotivoPro = mProps.IdMotivoPro
        !IdProcesado = mProps.IdProcesado
        !IdMarca = mProps.IdMarca
        !IdCliente = mProps.objGNComprobante.IdClienteRef
        !Motivo = mProps.objGNComprobante.NumTransCierrePOS
        If mProps.objGNComprobante.NumTransCierrePOS = 2 Then
            !BandVerifica = True
            !BandProceso = True
            If mProps.objGNComprobante.BandCierre = 1 Then
                !BandCarcasaISO = True
            End If
        End If
        !IdObra = mProps.objGNComprobante.IdObra
        .Update
        #If DAOLIB Then
                rs.Bookmark = rs.LastModified
        #End If
        
                
        .Move 0             'Para actualizar
        
        mProps.IdSecuencial = !Id             'Guarda el Id(Autonumerico)
        .Close
        
        'actualiza Tiempo Entrega
        sql = "Update Ivkardex set TiempoEntrega= " & mProps.IdSecuencial & " WHERE id=" & mProps.Id
        Set rsA = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    End With
End Sub

Public Property Let BandVerifica(value As Boolean)
    mProps.BandVerifica = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get BandVerifica() As Boolean
    BandVerifica = mProps.BandVerifica
End Property

Public Property Let BandProceso(value As Boolean)
    mProps.BandProceso = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get BandProceso() As Boolean
    BandProceso = mProps.BandProceso
End Property



Public Sub RecuperarIVInventarioDetalleISO(Optional ByVal TransIDIVInventarioDetalleISO As Long)
    Dim sql As String, rs As Recordset
    
    If TransIDIVInventarioDetalleISO = 0 Then TransIDIVInventarioDetalleISO = mProps.Id
    
    
    sql = " SELECT iviso.*, ivtam.codinventario as Codtamanio , ivtra.codinventario as Codtrabajo, ivmar.codinventario as Codmarca , "
    sql = sql & " ivmotcal.codmotivo as codmotivoCal ,ivmotpro.codmotivo as CodMOtivoPro, codprovcli as codcliente, ivdp.codinventario as CodProcesado, gin.numdocref as NumIngreso "
    sql = sql & " FROM IVInventarioDetalleISO iviso"
    sql = sql & " inner join gncomprobante gin on iviso.transid=gin.transid"
    sql = sql & " inner join ivinventario ivtam on iviso.idtamanio=ivtam.idinventario"
    sql = sql & " LEFT join ivinventario ivtra on iviso.idtrabajo=ivtra.idinventario"
    sql = sql & " inner join ivinventario ivmar on iviso.idmarca=ivmar.idinventario"
    sql = sql & " left join ivinventario ivdp on iviso.idProcesado=ivdp.idinventario"
    sql = sql & " left join motivo ivmotCal on iviso.idmotivoCal=ivmotcal.idmotivo"
    sql = sql & " left join motivo ivmotPro on iviso.idmotivoPro=ivmotPro.idmotivo"
    sql = sql & " left join pcprovcli pc on iviso.idcliente=pc.idprovcli"
    sql = sql & " Where ID = " & TransIDIVInventarioDetalleISO
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
    
    If rs.EOF Then Exit Sub
    With rs
        mProps.IdSecuencial = !Id             'Guarda el Id(Autonumerico)
        If Not IsNull(!Original) Then mProps.Original = !Original
        If Not IsNull(!Radial) Then mProps.Radial = !Radial
        If Not IsNull(!NumSerie) Then mProps.NumSerie = !NumSerie
        If Not IsNull(!DisenioOri) Then mProps.DisenioOri = !DisenioOri
        If Not IsNull(!IdTrabajo) Then mProps.IdTrabajo = !IdTrabajo
        If Not IsNull(!CodTrabajo) Then mProps.CodTrabajo = !CodTrabajo
        If Not IsNull(!IdTamanio) Then mProps.IdTamanio = !IdTamanio
        If Not IsNull(!CodTamanio) Then mProps.CodTamanio = !CodTamanio
'''        If Not IsNull(!secuencial) Then mProps.secuencial = !secuencial
        mProps.BandVerifica = !BandVerifica
        mProps.BandProceso = !BandProceso
        mProps.BandFactura = !BandFactura
        mProps.BandCarcasaISO = !BandCarcasaISO
        If Not IsNull(!IdMotivoCal) Then mProps.IdMotivoCal = !IdMotivoCal
        If Not IsNull(!CodMotivoCal) Then mProps.CodMotivoCal = !CodMotivoCal
        If Not IsNull(!IdMotivoPro) Then mProps.IdMotivoPro = !IdMotivoPro
        If Not IsNull(!CodMotivoPro) Then mProps.CodMotivoPro = !CodMotivoPro
        If Not IsNull(!IdMarca) Then mProps.IdMarca = !IdMarca
        If Not IsNull(!CodMarca) Then mProps.CodMarca = !CodMarca
        If Not IsNull(!IdCliente) Then mProps.IdCliente = !IdCliente
        If Not IsNull(!CodCliente) Then mProps.CodCliente = !CodCliente
        If Not IsNull(!IdProcesado) Then mProps.IdProcesado = !IdProcesado
        If Not IsNull(!CodProcesado) Then mProps.CodProcesado = !CodProcesado
        
        
        mProps.BandCarcasaUsada = !BandCarcasaUsada
        mProps.BandComprada = !BandComprada
        mProps.BandPagado = !BandPagado
        If Not IsNull(!Motivo) Then mProps.Motivo = !Motivo
        If Not IsNull(!ValorCarcasa) Then mProps.ValorCarcasa = !ValorCarcasa
        If Not IsNull(!ValorTrabajo) Then mProps.ValorTrabajo = !ValorTrabajo
        If Not IsNull(!NumIngreso) Then mProps.NumIngreso = !NumIngreso
        .Close
    End With
    Set rs = Nothing
End Sub

Public Property Let BandFactura(value As Boolean)
    mProps.BandFactura = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get BandFactura() As Boolean
    BandFactura = mProps.BandFactura
End Property

Public Property Let TransIDCalifica(ByVal value As Long)
    mProps.TransIDCalifica = value
End Property

Public Property Get TransIDCalifica() As Long
    TransIDCalifica = mProps.TransIDCalifica
End Property

Public Property Let TransIDProceso(ByVal value As Long)
    mProps.TransIDProceso = value
End Property

Public Property Get TransIDProceso() As Long
    TransIDProceso = mProps.TransIDProceso
End Property

Public Property Let TransIDFactura(ByVal value As Long)
    mProps.TransIDFactura = value
End Property

Public Property Get TransIDFactura() As Long
    TransIDFactura = mProps.TransIDFactura
End Property

Public Property Let BandCarcasaISO(value As Boolean)
    mProps.BandCarcasaISO = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get BandCarcasaISO() As Boolean
    BandCarcasaISO = mProps.BandCarcasaISO
End Property


Public Property Let IdMotivoCal(ByVal value As Long)
    mProps.IdMotivoCal = value
End Property

Public Property Get IdMotivoCal() As Long
    IdMotivoCal = mProps.IdMotivoCal
End Property

Public Property Let CodMotivoCal(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodMotivoCal Then
        If Len(value) > 0 Then
            'Actualiza IdInventario también
            sql = "SELECT idMotivo, CodMotivo " & _
                  "FROM Motivo WHERE CodMotivo='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IvKardex_CodMotivoCal", MSGERR_NOHAYCODIGO
            Else
                mProps.IdMotivoCal = rs!IdMotivo
            End If
            rs.Close
        Else
            mProps.IdMotivoCal = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodMotivoCal = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodMotivoCal() As String
    CodMotivoCal = Trim$(mProps.CodMotivoCal)
End Property


Public Property Let IdMotivoPro(ByVal value As Long)
    mProps.IdMotivoPro = value
End Property

Public Property Get IdMotivoPro() As Long
    IdMotivoPro = mProps.IdMotivoPro
End Property

Public Property Let CodMotivoPro(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodMotivoPro Then
        If Len(value) > 0 Then
            'Actualiza IdInventario también
            sql = "SELECT idMotivo, CodMotivo " & _
                  "FROM Motivo WHERE CodMotivo='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IvKardex_CodMotivoPro", MSGERR_NOHAYCODIGO
            Else
                mProps.IdMotivoPro = rs!IdMotivo
            End If
            rs.Close
        Else
            mProps.IdMotivoPro = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodMotivoPro = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodMotivoPro() As String
    CodMotivoPro = Trim$(mProps.CodMotivoPro)
End Property

Public Property Let IdMarca(ByVal value As Long)
    mProps.IdMarca = value
End Property

Public Property Get IdMarca() As Long
    IdMarca = mProps.IdMarca
End Property

Public Property Let CodMarca(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodMarca Then
        If Len(value) > 0 Then
            'Actualiza IdInventario también
            sql = "SELECT idInventario, CodInventario " & _
                  "FROM IVInventario WHERE CodInventario='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IvKardex_CodMarca", MSGERR_NOHAYCODIGO
            Else
                mProps.CodMarca = rs!CodInventario
                mProps.IdMarca = rs!IdInventario
            End If
            rs.Close
        Else
            mProps.IdPadreSub = 0
            mProps.CodMarca = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodMarca = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodMarca() As String
    CodMarca = Trim$(mProps.CodMarca)
End Property


Public Property Let IdCliente(ByVal value As Long)
    mProps.IdCliente = value
End Property

Public Property Get IdCliente() As Long
    IdCliente = mProps.IdCliente
End Property

Public Property Let CodCliente(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCliente Then
        If Len(value) > 0 Then
            'Actualiza IdInventario también
            sql = "SELECT idInventario, CodInventario " & _
                  "FROM IVInventario WHERE CodInventario='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IvKardex_CodCliente", MSGERR_NOHAYCODIGO
            Else
                mProps.CodCliente = rs!CodInventario
                mProps.IdCliente = rs!IdInventario
            End If
            rs.Close
        Else
            mProps.IdPadreSub = 0
            mProps.CodCliente = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodCliente = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodCliente() As String
    CodCliente = Trim$(mProps.CodCliente)
End Property


Public Property Get TransIDFacturaCarc() As Long
    TransIDFactura = mProps.TransIDFacturaCarc
End Property

Public Property Let TransIDFacturaCarc(value As Long)
    mProps.TransIDFacturaCarc = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Let BandCarcasaUsada(value As Boolean)
    mProps.BandCarcasaUsada = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get BandCarcasaUsada() As Boolean
    BandCarcasaUsada = mProps.BandCarcasaUsada
End Property

Public Property Let BandComprada(value As Boolean)
    mProps.BandComprada = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get BandComprada() As Boolean
    BandComprada = mProps.BandComprada
End Property

Public Property Let Motivo(ByVal value As Byte)
    mProps.Motivo = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Motivo() As Byte
    Motivo = mProps.Motivo
End Property

Public Function ObtieneNumeroPrecio(codItem As String) As Integer
    Dim iv As IVInventario, i As Integer, Precio As Currency
    Set iv = mProps.objGNComprobante.Empresa.RecuperaIVInventario(mProps.CodInventario)
     If Not (iv Is Nothing) Then
     
            If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "LIRA") > 0 Then
                'redondea a 2 decimales porque ellos tiene precios con 4 decimales pero el formato PT DEBE SER 2 DECIMALES
                Precio = Round(mProps.PrecioTotal / mProps.Cantidad / mProps.objGNComprobante.Cotizacion(""), 2)
                For i = 1 To 5
                    If Precio = Round(iv.Precio(i), 2) Then
                        ObtieneNumeroPrecio = i
                        Exit For
                    End If
                Next i
            Else
                Precio = Format(mProps.PrecioTotal / mProps.Cantidad / mProps.objGNComprobante.Cotizacion(""), Me.GNComprobante.FormatoMoneda)
                 For i = 1 To 5
                    If Precio = Format(iv.Precio(i), Me.GNComprobante.FormatoMoneda) Then
                        ObtieneNumeroPrecio = i
                        Exit For
                    End If
                Next i
            End If
    End If
    Set iv = Nothing
End Function

'Precio para el Numero de dias
Public Property Get PrecioND() As Currency
    If mProps.Cantidad Then
        If mProps.NumDias Then
            PrecioND = mProps.PrecioTotal / mProps.Cantidad / mProps.NumDias / mProps.objGNComprobante.Cotizacion("")
        Else
            PrecioND = mProps.PrecioTotal / mProps.Cantidad / mProps.objGNComprobante.Cotizacion("")
        End If
    End If
End Property

Public Property Let FechaDevol(value As Date)
    If mProps.FechaDevol <> value Then
        mProps.FechaDevol = value
    End If
End Property

Public Property Get FechaDevol() As Date
    FechaDevol = mProps.FechaDevol
End Property

Public Property Let NumDias(ByVal value As Currency)
    mProps.NumDias = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get NumDias() As Currency
    NumDias = mProps.NumDias
End Property

Public Property Let CantidadDev(ByVal value As Currency)
    mProps.CantidadDev = value
End Property

Public Property Get CantidadDev() As Currency
    CantidadDev = mProps.CantidadDev
End Property

Public Property Let FechaLleva(value As Date)
    If mProps.FechaLleva <> value Then
        mProps.FechaLleva = value
    End If
End Property

Public Property Get FechaLleva() As Date
    FechaLleva = mProps.FechaLleva
End Property

Public Property Let IdProcesado(ByVal value As Long)
    mProps.IdProcesado = value
End Property

Public Property Get IdProcesado() As Long
    IdProcesado = mProps.IdProcesado
End Property

Public Property Let CodProcesado(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodProcesado Then
        If Len(value) > 0 Then
            'Actualiza IdInventario también
            sql = "SELECT idInventario, CodInventario " & _
                  "FROM IVInventario WHERE CodInventario='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IvKardex_CodProcesado", MSGERR_NOHAYCODIGO
            Else
                mProps.IdPadreSub = rs!IdInventario
                mProps.CodProcesado = rs!CodInventario
                mProps.IdProcesado = rs!IdInventario
            End If
            rs.Close
        Else
            mProps.IdPadreSub = 0
            mProps.CodProcesado = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodProcesado = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodProcesado() As String
    CodProcesado = Trim$(mProps.CodProcesado)
End Property

Public Property Get TransIDCompra() As Long
    TransIDCompra = mProps.TransIDCompra
End Property

Public Property Let TransIDCompra(value As Long)
    mProps.TransIDCompra = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get ValorCarcasa() As Currency
    ValorCarcasa = mProps.ValorCarcasa
End Property

Public Property Let ValorCarcasa(value As Currency)
     mProps.ValorCarcasa = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get ValorTrabajo() As Currency
    ValorTrabajo = mProps.ValorTrabajo
End Property

Public Property Let ValorTrabajo(value As Currency)
    mProps.ValorTrabajo = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Let BandPagado(value As Boolean)
    mProps.BandPagado = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get BandPagado() As Boolean
    BandPagado = mProps.BandPagado
End Property


Public Property Get NumIngreso() As String
    NumIngreso = Trim$(mProps.NumIngreso)
End Property

Public Property Let NumIngreso(ByVal value As String)
    value = RTrim$(value)
    If Len(value) > 20 Then                 '*** MAKOTO 16/oct/00 Modifcado 40-->80
        Err.Raise ERR_INVALIDO, "IVKardex_NumIngreso", MSGERR_INVALIDO
    End If
    
    mProps.NumIngreso = value
    mProps.objGNComprobante.SetModificado
    
End Property


Public Sub GrabaConsumoDetalle()
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As IVConsumoDetalle, sql As String, band As Boolean
    Dim NumHijos As Integer, CountNumHijos As Integer, item As IVInventario
    Dim codRelleno As String, codCemento As String, codCojin As String
    Dim porRelleno As Currency, porCemento As Currency, porCojin As Currency
    Dim v As Variant, ITEMCONS As IVInventario, Banda As IVInventario
    Dim costoRelleno As Currency, costoCemento As Currency, costoCojin As Currency, costobanda As Currency, codbanda As String
    On Error GoTo CapturaError
                If Len(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("Porcentaje_RELLENO")) > 0 Then
                    v = Split(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("Porcentaje_RELLENO"), ",")
                    codRelleno = v(0)
                    porRelleno = Round((v(1) / 100), 4)
                    Set ITEMCONS = mProps.objGNComprobante.Empresa.RecuperaIVInventarioQuick(codRelleno)
                    costoRelleno = 0
                    If Not ITEMCONS Is Nothing Then
                        costoRelleno = ITEMCONS.CostoDouble2(mProps.objGNComprobante.FechaTrans, _
                                    1, _
                                    mProps.objGNComprobante.transid, _
                                    mProps.objGNComprobante.HoraTrans)  '*** MAKOTO 08/dic/00 Agregado Hora
                    End If
                End If
                
                If Len(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("Porcentaje_CEMENTO")) > 0 Then
                    v = Split(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("Porcentaje_CEMENTO"), ",")
                    codCemento = v(0)
                    porCemento = Round((v(1) / 100), 4)
                    Set ITEMCONS = mProps.objGNComprobante.Empresa.RecuperaIVInventarioQuick(codCemento)
                    costoCemento = 0
                    If Not ITEMCONS Is Nothing Then
                        costoCemento = ITEMCONS.CostoDouble2(mProps.objGNComprobante.FechaTrans, _
                                    1, _
                                    mProps.objGNComprobante.transid, _
                                    mProps.objGNComprobante.HoraTrans)  '*** MAKOTO 08/dic/00 Agregado Hora
                    End If
                End If
                
                If Len(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("Porcentaje_COJIN")) > 0 Then
                    v = Split(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("Porcentaje_COJIN"), ",")
                    codCojin = v(0)
                    porCojin = Round((v(1) / 100), 4)
                    Set ITEMCONS = mProps.objGNComprobante.Empresa.RecuperaIVInventarioQuick(codCojin)

                    costoCojin = 0
                    If Not ITEMCONS Is Nothing Then
                        costoCojin = ITEMCONS.CostoDouble2(mProps.objGNComprobante.FechaTrans, _
                                    1, _
                                    mProps.objGNComprobante.transid, _
                                    mProps.objGNComprobante.HoraTrans)  '*** MAKOTO 08/dic/00 Agregado Hora
                    End If
                End If
        For i = 1 To 3
            numReg = Me.AddConsumoDetalle
            Set obj = mcolConsumoDetalle.item(i)
            Select Case i
                Case 1
                    obj.Cantidad = Me.Cantidad * porRelleno * -1
                    obj.CodInventario = codRelleno
                    obj.Costo = costoRelleno * obj.Cantidad
                Case 2
                    obj.Cantidad = Me.Cantidad * porCemento * -1
                    obj.CodInventario = codCemento
                    obj.Costo = costoCemento * obj.Cantidad
                Case 3
                    obj.Cantidad = Me.Cantidad * porCojin * -1
                    obj.CodInventario = codCojin
                    obj.Costo = costoCojin * obj.Cantidad
                End Select
                obj.Orden = i
        Next i
    
        'Banda
            numReg = Me.AddConsumoDetalle
            Set obj = mcolConsumoDetalle.item(i)
            Set Banda = mProps.objGNComprobante.Empresa.RecuperaIVInventario(Me.CodInventario)
            costobanda = Banda.CostoDouble2(mProps.objGNComprobante.FechaTrans, 1, mProps.objGNComprobante.transid, mProps.objGNComprobante.HoraTrans)
    
            obj.Cantidad = Me.Cantidad * -1
            obj.CodInventario = Banda.CodInventario
            obj.Costo = costobanda * obj.Cantidad
            obj.Orden = 4
    
    
    'Elimina  todos  los registros
    sql = "Delete Ivcd " & _
          "From IvKardex Ivk " & _
          "INNER JOIN IVConsumoDetalle Ivcd ON IVk.Id = Ivcd.IdKardexRef "
          sql = sql & " Where Ivcd.IdkardexRef = " & mProps.Id

    mProps.objGNComprobante.Empresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    band = True
    If Not mcolConsumoDetalle Is Nothing Then
        sql = "SELECT * FROM IVConsumoDetalle WHERE 1=0"
        Set rs1 = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
        NumHijos = mcolConsumoDetalle.Count
        CountNumHijos = 0
        For i = 1 To mcolConsumoDetalle.Count
'            rs1.AddNew /jeaa 28-01-04
            If CountNumHijos < NumHijos Then
                CountNumHijos = CountNumHijos + 1
                Set obj = mcolConsumoDetalle.item(i)
                If obj.IdInventario = 0 Then
                        obj.Recuperar obj.CodInventario
                        Set item = mProps.objGNComprobante.Empresa.RecuperaIVInventario(obj.CodInventario)
                        If verificacionDuiplicado(obj.IdInventario) Then
                            RemoveConsumoDetalle i
                            i = i - 1
                        End If
                End If
                If band Then
                    rs1.AddNew
                    rs1!transid = mProps.objGNComprobante.transid
                    rs1!IdInventario = item.IdInventario
                    rs1!idkardexref = mProps.Id
                    rs1!cant = obj.Cantidad
                    rs1!Precio = obj.Precio
                    rs1!Orden = obj.Orden
                    rs1!Costo = obj.Costo
                    rs1!Ticket = mProps.TiempoEntrega
                    rs1.Update
                End If
            End If
            band = True
            Set item = Nothing
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


Public Property Get NumFamiliaDetalle() As Long
    If mcolConsumoDetalle Is Nothing Then
        NumFamiliaDetalle = 0
    Else
        NumFamiliaDetalle = mcolConsumoDetalle.Count
    End If
    'If ObtieneUltimo(False) Then PrecioUltimoEgreso = mUEgreso.Precio
End Property

Private Sub CargaConsumoDetalle()
    Dim rs As Recordset, sql As String, obj As IVConsumoDetalle
    Dim i As Long
    Set mcolConsumoDetalle = New Collection
        sql = "Select IV.IdInventario, Iv.CodInventario, Iv.Descripcion, "
        sql = sql & " Ivcd.Cant , Ivcd.Precio, Ivcd.orden "
        sql = sql & " From  ivkardex IVk "
        sql = sql & " INNER JOIN IVConsumoDetalle Ivcd "
        sql = sql & " INNER JOIN IVInventario Iv "
        sql = sql & " ON IV.IdInventario = Ivcd.IdInventario "
        sql = sql & " ON IVk.Id = Ivcd.Idkardexref"
        sql = sql & " Where Ivcd.Idkardexref = " & mProps.Id
        sql = sql & " order by Ivcd.orden"
    
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddConsumoDetalle
            Set obj = mcolConsumoDetalle.item(i)
                  obj.IdInventario = !IdInventario
                  obj.CodInventario = !CodInventario
                  obj.Descripcion = !Descripcion
                 obj.Cantidad = !cant
                obj.Precio = !Precio
                obj.Orden = !Orden
           .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
End Sub

Public Function RecuperaConsumoDetalle(ix As Long) As IVConsumoDetalle
    Set RecuperaConsumoDetalle = mcolConsumoDetalle.item(ix)
End Function


Public Function AddConsumoDetalle() As Long
    Dim nuevo As IVConsumoDetalle
    
    If mcolConsumoDetalle Is Nothing Then CargaConsumoDetalle
    Set nuevo = New IVConsumoDetalle
    Set nuevo.IVKardex = Me  'Me  olvide  del padre
    'Proporciona valores predeterminados
    nuevo.Cantidad = 1
    
    mcolConsumoDetalle.Add nuevo
    AddConsumoDetalle = mcolConsumoDetalle.Count    'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    
   
End Function

Public Sub RemoveConsumoDetalle(ByVal ix As Long, Optional obj As IVConsumoDetalle)
    If mcolConsumoDetalle Is Nothing Then CargaConsumoDetalle
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolConsumoDetalle.Count
            If mcolConsumoDetalle.item(ix) Is obj Then
                mcolConsumoDetalle.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolConsumoDetalle.Remove ix
    End If
    
End Sub

Private Function verificacionDuiplicado(ByVal IdInventario As String) As Boolean
    Dim rs As Recordset, sql As String
    verificacionDuiplicado = False
    'Recupera los detalles de existencia
    sql = "SELECT IdInventario FROM ivmateria where idinventario =" & IdInventario
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        verificacionDuiplicado = True
    End If
    Set rs = Nothing
End Function

Public Property Let Nota1(ByVal value As String)
value = RTrim$(value)
    If Len(value) > 250 Then                '*** MAKOTO 16/oct/00 Modifcado 40-->80                                            'auc modificado 12/08/08 de 80 a 200
        value = Left(RTrim$(value), 250)
        Err.Raise ERR_INVALIDO, "IVKardex_Nota", MSGERR_INVALIDO & " " & "Maximo 250 chr"
    End If
   mProps.Nota1 = value
    mProps.objGNComprobante.SetModificado

End Property

Public Property Get Nota1() As String
    Nota1 = RTrim$(mProps.Nota1)
End Property


