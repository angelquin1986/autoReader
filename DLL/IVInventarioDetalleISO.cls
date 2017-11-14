VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVInventarioDetalleISO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROP
    id As Long
    TransIDIngreso As Long
    IdInventario As Long
    CodInventario As String
    IDTamanio As Long
    CodTamanio As String
    Original As String
    Radial As String
    NumSerie As String
    DisenioOri As String
    IDTrabajo As Long
    CodTrabajo As String
    
    TransIDCalifica As Long
    IDMotivoCalifica As Long
    CodMotivoCalifica As String
    
    TransIDProceso As Long
    IDMotivoProceso As Long
    CodMotivoProceso As String
    
    secuencial As Long
    BandVerifica As Boolean
    BandProceso As Boolean
End Type

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
'    Set item = mProps.objGNComprobante.Empresa.RecuperaIVInventario(mProps.IdInventario)
       
           
     If Not mProps.objGNComprobante.GNTrans.MostrarSubItemsImp Then  'AUC cambiado para arreglar la secuencia
        Set item = mProps.objGNComprobante.Empresa.RecuperaIVInventario(mProps.CodInventario)
        If item.tipo = Preparacion Then
            'mProps.Cantidad = -1
            mProps.Cantidad = Me.Cantidad
        End If
    End If
       'La cantidad no puede ser 0.
    Set item = mProps.objGNComprobante.Empresa.RecuperaIVInventario(mProps.CodInventario)
    If item.tipo <> rubro Then
        If mProps.Cantidad = 0 Then
           Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", _
              "La cantidad no puede ser 0. (" & mProps.CodInventario & ")"
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
    If mProps.objGNComprobante.GNTrans.IVControlPrecio0 Then
        If mProps.PrecioRealTotal = 0 Then
            Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", _
                "En ésta transacción no se puede grabar con precio 0. (" & mProps.CodInventario & ")"
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
                        'Err.Raise ERR_INVALIDO, "IVKardex_VerificaDatos", msg
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
    
    If Me.GNComprobante.EsNuevo Then
   
        If Len(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("SecuencialItems")) > 0 Then
            SecuencialItem = CDbl(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("SecuencialItems")) + 1
        End If
    End If

    
    'Asegura que en IVExist haya un registro correspondiente
    AseguraBodega
    
    '***Agregado. 02/sep/2003
    'Busca Numero de Precio seleccionado
    If Me.GNComprobante.EsNuevo Then
        If mProps.idpadre <> 0 Then
            BuscaNumeroPrecioPadre mProps.idpadre
        Else
            BuscaNumeroPrecio
        End If
    
        If Len(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("SecuencialItems")) > 0 Then
            SecuencialItem = CDbl(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("SecuencialItems")) + 1
        End If
    End If
    sql = "SELECT * FROM IVKardex WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !TransID = mProps.objGNComprobante.TransID
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
        !NumeroPrecio = mProps.NumeroPrecio
        !ValorRecargoItem = mProps.ValorRecargoItem '***Agregado. Angel. 29/jul/2004
        
        
        If Me.GNComprobante.EsNuevo And Me.GNComprobante.GNTrans.IVSecuencialItem Then
            !TiempoEntrega = SecuencialItem
            mProps.secuencial = SecuencialItem
        Else
            !TiempoEntrega = mProps.TiempoEntrega   ' jeaa 22/09/2005
        End If
        !IdICE = mProps.IdICE '***Agregado. jeaa 21/07/2006
        !ValorICEItem = mProps.PorcentajeRecargoIce '***Agregado. jeaa 21/07/2006
        !bandImprimir = mProps.bandImprimir 'AUC 21/07/06
        !idpadre = mProps.idpadre 'AUC 15/09/06
        !bandVer = mProps.bandVer 'AUC 21/07/06
        '!bandVer = mProps.bandVer 'AUC 21/07/06
        !IdPadreSub = mProps.IdPadreSub  'AUC 15/09/06

        .Update
        
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.id = !id             'Guarda el Id(Autonumerico)
        .Close
        'AUC 23/03/06 agregado bandera para generar o no el kardex de subitems
        With mProps.objGNComprobante.GNTrans
            Set item = Me.GNComprobante.Empresa.RecuperaIVInventario(Me.IdInventario)
             'If item.tipo = 4 And .IVTipoTrans = "E" And .GenerarKardexSubItem Then
              If item.tipo = 4 And .IVTipoTrans = "E" And .MostrarSubItemsImp Then
             
                 GeneraPreparacion item 'AUC 01/15/06  para retroburguer
             End If
            '''antes If .IVTipoTrans = "I" Then
            If mProps.objGNComprobante.GNTrans.IVActualizaCostoUCenItems And mProps.objGNComprobante.EsNuevo Then
                'mProps.objGNComprobante.Empresa.GrabarCostoUltIngreso mProps.IdInventario, mProps.CostoTotal / mProps.Cantidad 'AUC 14/12/07 cambiado AUC debe ser costoreal
                 mProps.objGNComprobante.Empresa.GrabarCostoUltIngreso mProps.IdInventario, mProps.CostoRealTotal / mProps.Cantidad
            End If

           If .IVTipoTrans = "A" Or .IVTipoTrans = "E" And mProps.objGNComprobante.GNTrans.CodPantalla = "IV_RESALQ" Then
                 GrabarIVReservacion
            End If
        End With
        
        If mProps.objGNComprobante.GNTrans.IVSecuencialItem Then
            GrabarIVInventarioDetalleISO
        End If
        If mProps.objGNComprobante.GNTrans.CodPantalla = "IVBDISO" And mProps.Cantidad < 0 Then
            mProps.objGNComprobante.ActualizaBandVerifica mProps.TiempoEntrega, mProps.bandVer
        End If
        
        If mProps.objGNComprobante.GNTrans.CodPantalla = "IVISOPRD" Then
            mProps.objGNComprobante.ActualizaBandProceso mProps.TiempoEntrega, mProps.bandImprimir
        End If
        
        If Len(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("SecuencialItems")) > 0 Then
            mProps.objGNComprobante.Empresa.GNOpcion.AsignarValor "SecuencialItems", SecuencialItem
            mProps.objGNComprobante.Empresa.GNOpcion.Grabar
        End If
        
    End With
    Set item = Nothing
    Set rs = Nothing
    Set pc = Nothing
End Sub

Friend Sub Recuperar(rs As Recordset)
    With rs
        If Not .EOF Then
            mProps.id = !id
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
            If Not IsNull(!NumeroPrecio) Then mProps.NumeroPrecio = !NumeroPrecio '***Agregado. 02/sep/2003. Angel
            If Not IsNull(!ValorRecargoItem) Then mProps.ValorRecargoItem = !ValorRecargoItem '***Agregado. 29/jul/2004. Angel
            
            If Not IsNull(!IdCuentaActivo) Then mProps.auxIdCuentaActivo = !IdCuentaActivo
            If Not IsNull(!IdCuentaCosto) Then mProps.auxIdCuentaCosto = !IdCuentaCosto
            If Not IsNull(!IdCuentaVenta) Then mProps.auxIdCuentaVenta = !IdCuentaVenta
            If Not IsNull(!TiempoEntrega) Then mProps.TiempoEntrega = !TiempoEntrega  ' *** jeaa 22/09/2005
            If Not IsNull(!IdICE) Then mProps.IdICE = !IdICE '***Agregado. jeaa 21/07/2006
            If Not IsNull(!ValorICEItem) Then mProps.PorcentajeRecargoIce = !ValorICEItem
            If Not IsNull(!idpadre) Then mProps.idpadre = !idpadre
            If Not IsNull(!bandImprimir) Then mProps.bandImprimir = !bandImprimir  ' *** AUC 08/06/2006
            If Not IsNull(!IdPadreSub) Then mProps.IdPadreSub = !IdPadreSub     ' *** AUC 16/07/2007
            If Not IsNull(!bandVer) Then mProps.bandVer = !bandVer   ' *** AUC 27/06/2006
            If Not IsNull(!codItemPadre) Then mProps.codItemPadre = !codItemPadre
            If Not IsNull(!CodItemPadreSub) Then mProps.CodItemPadreSub = !CodItemPadreSub
            If mProps.objGNComprobante.GNTrans.IVSecuencialItem Then
                RecuperarIVInventarioDetalleISO mProps.id
            End If
          
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
    mProps.id = 0
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Set GNComprobante(value As GNComprobante)
    Set mProps.objGNComprobante = value
    
End Property

Public Property Get GNComprobante() As GNComprobante
    Set GNComprobante = mProps.objGNComprobante
End Property

'AUC cambiado para el reproceso de recetas 19/10/06
Public Property Get id() As Long
    id = mProps.id
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
            sql = "SELECT IdInventario,IdCuentaActivo,IdCuentaCosto,IdCuentaVenta, ValorRecargo, isnull(idice,0) as idice, isnull(porcentajeice,0) as porcentajeice " & _
                  "FROM IVInventario WHERE CodInventario='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "CTLibroDetalle_CodInventario", MSGERR_NOHAYCODIGO
            Else
                mProps.IdInventario = rs!IdInventario
                mProps.auxIdCuentaActivo = rs!IdCuentaActivo
                mProps.auxIdCuentaCosto = rs!IdCuentaCosto
                mProps.auxIdCuentaVenta = rs!IdCuentaVenta
                mProps.ValorRecargoItem = rs!ValorRecargo '***Agregado. 29/jul/2004. Angel
                mProps.PorcentajeRecargoIce = rs!PorcentajeICE '***Agregado. 21/07/2006 jeaa
                mProps.IdICE = rs!IdICE '***Agregado. 21/07/2006 jeaa
            End If
            rs.Close
        Else
            mProps.IdInventario = 0
            mProps.auxIdCuentaActivo = 0
            mProps.auxIdCuentaCosto = 0
            mProps.auxIdCuentaVenta = 0
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
    If Len(value) > 80 Then                 '*** MAKOTO 16/oct/00 Modifcado 40-->80
        Err.Raise ERR_INVALIDO, "IVKardex_Nota", MSGERR_INVALIDO
    End If
    
    mProps.Nota = value
    mProps.objGNComprobante.SetModificado
End Property

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
        neto = Abs(Me.GNComprobante.IVKardexTotal(False))    'Total NETO sin recargo prorateado
        neto = neto - Me.GNComprobante.IVKardexDescItemTotal
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
    If (value < 0) Or (value > 5) Then value = 0
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
        !TransID = mProps.objGNComprobante.TransID
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
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
         mProps.id = !id             'Guarda el Id(Autonumerico)
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
                    If item.tipo = Preparacion Then    'Solo para Preparacion
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
                               If item1.tipo = Preparacion Then
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
        If item.tipo <> Preparacion Then        'Solo para Preparacion
            Me.GNComprobante.IVKardex(indice).IdInventario = item.RecuperaDetalleFamilia(k).IdInventario
            Set ItemPrepara = Me.GNComprobante.Empresa.RecuperaIVInventario(Me.GNComprobante.IVKardex(indice).IdInventario)
            If Not (ItemPrepara Is Nothing) Then
                If Not ItemPrepara.BandServicio Then
                    exist = ItemPrepara.Existencia(Me.GNComprobante.IVKardex(indice).CodBodega)
                    sumaCant = 0
                    'Obtiene la cantidad original en caso de modificación
                    cant_ori = Me.GNComprobante.VerificarExistencia_CantOrig(Me.id)
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
            If Not item.BandServicio And item.tipo <> Preparacion Then
                sumaCant = 0
                'Obtiene la cantidad original en caso de modificación
                cant_ori = Me.GNComprobante.VerificarExistencia_CantOrig(Me.id)
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
                sql = "SELECT IdICE, isnull(PorcentajeICE,0) as PorcentajeICE " & _
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
                    mProps.IdICE = rs!IDMotivo
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
    v = MiCCur(Format$(v, Me.GNComprobante.FormatoMoneda))
    
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
        !TransID = mProps.objGNComprobante.TransID
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


Public Property Let IDTamanio(ByVal value As Long)
    mProps.IDTamanio = value
End Property

Public Property Get IDTamanio() As Long
    IDTamanio = mProps.IDTamanio
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
                mProps.IDTamanio = rs!IdInventario
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

Public Property Let IDTrabajo(ByVal value As Long)
    mProps.IDTrabajo = value
End Property

Public Property Get IDTrabajo() As Long
    IDTrabajo = mProps.IDTrabajo
End Property

Public Property Let secuencial(ByVal value As Long)
    mProps.secuencial = value
End Property

Public Property Get secuencial() As Long
    secuencial = mProps.secuencial
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
                mProps.IDTrabajo = rs!IdInventario
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
    
   
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM IVInventarioDetalleISO WHERE IDKardex=" & mProps.id
        Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
        If Not rs.EOF Then
            #If DAOLIB Then
                rs.Edit
            #End If
        Else
            sql = "SELECT * FROM IVInventarioDetalleISO WHERE 1=0"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
            rs.AddNew
        End If
        
    With rs
        !IDKardex = mProps.id                       'Es el mismo asignado a GnComprobante
        !IDTamanio = mProps.IDTamanio
        !Original = mProps.Original
        !Radial = mProps.Radial
        !DisenioOri = mProps.DisenioOri
        !IDTrabajo = mProps.IDTrabajo
        !secuencial = mProps.secuencial
        !NumSerie = mProps.NumSerie
        !BandVerifica = mProps.bandVer
        !BandProceso = mProps.BandProceso
        !TransID = mProps.objGNComprobante.TransID
        .Update
        #If DAOLIB Then
                rs.Bookmark = rs.LastModified
        #End If
        .Move 0             'Para actualizar
        .Close
        
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
    
    If TransIDIVInventarioDetalleISO = 0 Then TransIDIVInventarioDetalleISO = mProps.id
    
    
    sql = " SELECT iviso.*, ivtam.codinventario as Codtamanio , ivtra.codinventario as Codtrabajo"
    sql = sql & " FROM IVInventarioDetalleISO iviso"
    sql = sql & " inner join ivinventario ivtam on iviso.idtamanio=ivtam.idinventario"
    sql = sql & " inner join ivinventario ivtra on iviso.idtrabajo=ivtra.idinventario"
    sql = sql & " Where IDKardex = " & TransIDIVInventarioDetalleISO
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
    
    If rs.EOF Then Exit Sub
    With rs
        If Not IsNull(!Original) Then mProps.Original = !Original
        If Not IsNull(!Radial) Then mProps.Radial = !Radial
        If Not IsNull(!NumSerie) Then mProps.NumSerie = !NumSerie
        If Not IsNull(!DisenioOri) Then mProps.DisenioOri = !DisenioOri
        If Not IsNull(!IDTrabajo) Then mProps.IDTrabajo = !IDTrabajo
        If Not IsNull(!IDTamanio) Then mProps.IDTamanio = !IDTamanio
        If Not IsNull(!CodTamanio) Then mProps.CodTamanio = !CodTamanio
        If Not IsNull(!CodTrabajo) Then mProps.CodTrabajo = !CodTrabajo
        If Not IsNull(!secuencial) Then mProps.secuencial = !secuencial
        If Not IsNull(!BandVerifica) Then mProps.BandVerifica = !BandVerifica
        If Not IsNull(!BandProceso) Then mProps.BandProceso = !BandProceso
        .Close
    End With
    Set rs = Nothing
End Sub





