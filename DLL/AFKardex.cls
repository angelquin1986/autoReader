VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AFKardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROP
    Id As Long
    IdInventario As Long
    CodInventario As String
    IdBodega As Long
    CodBodega As String
    Cantidad As Currency
    
    CostoTotal As Currency
    CostoRealTotal As Currency
    PrecioTotal As Currency
    PrecioRealTotal As Currency
    
    Descuento As Currency
    IVA As Currency
    Nota As String
    Orden As Integer
    
    auxIdCuentaActivo As Long
    auxIdCuentaCosto As Long
    auxIdCuentaVenta As Long
    auxIdCuentaDepreGasto As Long
    auxIdCuentaDepreAcumulada As Long
    auxIdCuentaRevaloriza As Long
    auxIdCuentaDepRevaloriza As Long
    objGNComprobante As GNComprobante           'Referencia al objeto padre
    
    NumeroPrecio As Byte            '***Agregado. 02/sep/2003. Angel
    TiempoEntrega As String '***Agregado. 22/09/2005. jeaa
    NumDepreReval As Currency
    FechaIniDepReval As Date
End Type

Private mProps As T_PROP

Private Sub AseguraBodega()
    Dim sql As String, rs As Recordset
    
    sql = "SELECT IdBodega FROM AFExist WHERE IdInventario=" & _
                mProps.IdInventario & " AND IdBodega=" & mProps.IdBodega
    
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
    'Si no existe el registro correspondiente, lo inserta
    If rs.EOF Then
        sql = "INSERT INTO AFExist (IdInventario,IdBodega,Exist) " & _
              "VALUES(" & mProps.IdInventario & "," & mProps.IdBodega & "," & 0 & ")"
        mProps.objGNComprobante.Empresa.Execute sql, True
    End If
    
    rs.Close
    
   
    Set rs = Nothing
End Sub

Public Sub VerificaDatos()
    Dim cant As Currency
    Dim item As AFInventario
    Set item = mProps.objGNComprobante.Empresa.RecuperaAFInventario(mProps.IdInventario)
       
       'La cantidad no puede ser 0.
    If mProps.objGNComprobante.CodTrans <> "RA" Then
       If mProps.Cantidad = 0 Then
            Err.Raise ERR_INVALIDO, "AFKardex_VerificaDatos", _
                "La cantidad no puede ser 0. (" & mProps.CodInventario & ")"
         End If
    End If
    
    If mProps.IdBodega = 0 Then
        Err.Raise ERR_INVALIDO, "AFKardex_VerificaDatos", _
            "Debe seleccionar una bodega. (" & mProps.CodInventario & ")"
    End If

    If mProps.IdInventario = 0 Then
        Err.Raise ERR_INVALIDO, "AFKardex_VerificaDatos", _
            "Debe seleccionar un item."
    End If
    
    '*** MAKOTO 07/feb/01 Agregado
    'Si la transacción está configurado que NO permita grabar con costo 0
    If mProps.objGNComprobante.GNTrans.IVNoPermitirCostoCero And Not item.BandServicio Then
        If mProps.CostoRealTotal = 0 Then
            Err.Raise ERR_INVALIDO, "AFKardex_VerificaDatos", _
                "En ésta transacción no se puede grabar con costo 0. (" & mProps.CodInventario & ")"
        End If
    End If
    '*** Oliver 15/07/2003 Agregado
    'Si la transaccion esta configurada para controlar que no graben con percios 0 en los items
    If mProps.objGNComprobante.GNTrans.IVControlPrecio0 Then
        If mProps.PrecioRealTotal = 0 Then
            Err.Raise ERR_INVALIDO, "AFKardex_VerificaDatos", _
                "En ésta transacción no se puede grabar con precio 0. (" & mProps.CodInventario & ")"
        End If
    End If
    'Modificado  DIEGO 09/06/2001
    Dim cantOtras As Currency, msg As String, CantLimite As Currency, Saldo As Currency
    Dim fmt As String
    CantLimite = 0
    cantOtras = 0

            'control para HORMIPISOS
    If item.BandServicio Then Exit Sub
            
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
                        Err.Raise ERR_INVALIDO, "AFKardex_VerificaDatos", msg
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
                        'Err.Raise ERR_INVALIDO, "AFKardex_VerificaDatos", msg
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
                            Err.Raise ERR_INVALIDO, "AFKardex_VerificaDatos", msg
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
                            'Err.Raise ERR_INVALIDO, "AFKardex_VerificaDatos", msg
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
                Err.Raise ERR_INVALIDO, "AFKardex_VerificaDatos", _
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
                    Err.Raise ERR_INVALIDO, "AFKardex_VerificaDatos", msg
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
                                Saldo = cantGrabada + CantLimite
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
                                Err.Raise ERR_LIMITEITEM, "AFKardex.VerificaDatos", msg
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
                                Err.Raise ERR_LIMITEITEM, "AFKardex.VerificaDatos", msg
                            Else
                                'sino no genera error sólo saca mensaje
                                MsgBox msg, vbInformation
                            End If
                        End If
                    Else
                    mProps.objGNComprobante.VerificarLimiteItem Me.CodInventario, CantLimite, cantGrabada

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
                                Err.Raise ERR_LIMITEITEM, "AFKardex.VerificaDatos", msg
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
'    sql = "SELECT  Sum(Cantidad) as Total   FROM AFKardex Where TransID = " & _
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
    Dim i As Long, ivk As AFKardex, cantOtras As Currency
    
    For i = 1 To mProps.objGNComprobante.CountIVKardex
        Set ivk = mProps.objGNComprobante.AFKardex(i)
        If (Not (ivk Is Me)) And (ivk.CodInventario = Me.CodInventario) Then
            cantOtras = cantOtras + ivk.Cantidad
        End If
    Next i
    Set ivk = Nothing

    SumaItemRepetido = cantOtras
End Function

Friend Sub Grabar()
    Dim sql As String, rs As Recordset
    Dim item As AFInventario
    
    VerificaDatos
    
    'Asegura que en IVExist haya un registro correspondiente
    AseguraBodega
    
    '***Agregado. 02/sep/2003
    'Busca Numero de Precio seleccionado
'    If Me.GNComprobante.EsNuevo Then BuscaNumeroPrecio
    
    sql = "SELECT * FROM AFKardex WHERE 1=0"
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
        !NumeroPrecio = mProps.NumeroPrecio
        !TiempoEntrega = mProps.TiempoEntrega
        !NumDepreReval = mProps.NumDepreReval
        !FechaIniDepReval = mProps.FechaIniDepReval
        .Update
        
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.Id = !Id             'Guarda el Id(Autonumerico)
        .Close
        'AUC 23/03/06 agregado bandera para generar o no el kardex de subitems
        With mProps.objGNComprobante.GNTrans
            If .IVActualizaCostoRefItems Then
                 mProps.objGNComprobante.Empresa.GrabarCostoAFUltIngreso mProps.IdInventario, mProps.CostoRealTotal / mProps.Cantidad, _
                mProps.objGNComprobante.IdProveedorRef, mProps.objGNComprobante.FechaTrans, mProps.objGNComprobante.NumSerie + mProps.objGNComprobante.NumSecuencial
            End If
        End With
        
        sql = "SELECT * FROM AFKardexCustodio WHERE idinventario=" & mProps.IdInventario
        Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
        If rs.RecordCount = 0 Then
        End If
        
        
        
    End With
    Set item = Nothing
    Set rs = Nothing

End Sub

Friend Sub Recuperar(rs As Recordset)
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
            If Not IsNull(!NumeroPrecio) Then mProps.NumeroPrecio = !NumeroPrecio '***Agregado. 02/sep/2003. Angel
            
            If Not IsNull(!IdCuentaActivo) Then mProps.auxIdCuentaActivo = !IdCuentaActivo
            If Not IsNull(!IdCuentaCosto) Then mProps.auxIdCuentaCosto = !IdCuentaCosto
            If Not IsNull(!IdCuentaVenta) Then mProps.auxIdCuentaVenta = !IdCuentaVenta
            If Not IsNull(!IdCuentaDepreGasto) Then mProps.auxIdCuentaDepreGasto = !IdCuentaDepreGasto
            If Not IsNull(!IdCuentaDepreAcumulada) Then mProps.auxIdCuentaDepreAcumulada = !IdCuentaDepreAcumulada
            If Not IsNull(!IdCuentaRevaloriza) Then mProps.auxIdCuentaRevaloriza = !IdCuentaRevaloriza
            If Not IsNull(!IdCuentaDepRevaloriza) Then mProps.auxIdCuentaDepRevaloriza = !IdCuentaDepRevaloriza
            
            
            If Not IsNull(!TiempoEntrega) Then mProps.TiempoEntrega = !TiempoEntrega  ' *** jeaa 22/09/2005
            
            If Not IsNull(!NumDepreReval) Then mProps.NumDepreReval = !NumDepreReval
            If Not IsNull(!FechaIniDepReval) Then mProps.FechaIniDepReval = !FechaIniDepReval
            
        End If
    End With
End Sub

Friend Function Clone() As AFKardex
    Dim nuevo As AFKardex
    
    Set nuevo = New AFKardex
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

Friend Property Get Id() As Long
    Id = mProps.Id
End Property


Public Property Let IdInventario(ByVal value As Long)
    mProps.IdInventario = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get IdInventario() As Long
    IdInventario = mProps.IdInventario
End Property


Public Property Let CodInventario(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodInventario Then
        If Len(value) > 0 Then
            'Actualiza IdInventario también
            sql = "SELECT IdInventario,IdCuentaActivo,IdCuentaCosto,IdCuentaVenta, "
            sql = sql & " IdCuentaDepreGasto, IdCuentaDepreAcumulada, IdCuentaRevaloriza, IdCuentaDepRevaloriza "
            sql = sql & " FROM AFInventario WHERE CodInventario='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "CTLibroDetalle_CodInventario", MSGERR_NOHAYCODIGO
            Else
                mProps.IdInventario = rs!IdInventario
                mProps.auxIdCuentaActivo = rs!IdCuentaActivo
                mProps.auxIdCuentaCosto = rs!IdCuentaCosto
                mProps.auxIdCuentaVenta = rs!IdCuentaVenta
                mProps.auxIdCuentaDepreGasto = rs!IdCuentaDepreGasto
                mProps.auxIdCuentaDepreAcumulada = rs!IdCuentaDepreAcumulada
                mProps.auxIdCuentaRevaloriza = rs!IdCuentaRevaloriza
                mProps.auxIdCuentaDepRevaloriza = rs!IdCuentaDepRevaloriza
            End If
            rs.Close
        Else
            mProps.IdInventario = 0
            mProps.auxIdCuentaActivo = 0
            mProps.auxIdCuentaCosto = 0
            mProps.auxIdCuentaVenta = 0
            mProps.auxIdCuentaDepreGasto = 0
            mProps.auxIdCuentaDepreAcumulada = 0
            mProps.auxIdCuentaRevaloriza = 0
            mProps.auxIdCuentaDepRevaloriza = 0

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
            sql = "SELECT IdBodega FROM AFBodega WHERE CodBodega='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "AFKArdex_CodBodega", MSGERR_NOHAYCODIGO
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
        Err.Raise ERR_INVALIDO, "AFKardex_Nota", MSGERR_INVALIDO
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
        Err.Raise ERR_INVALIDO, "AFKardex_TiempoEntrega", MSGERR_INVALIDO
    End If
    
    mProps.TiempoEntrega = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get TiempoEntrega() As String
    TiempoEntrega = RTrim$(mProps.TiempoEntrega)
End Property

Public Property Get auxIdCuentaActivo() As Long
    auxIdCuentaActivo = mProps.auxIdCuentaActivo
End Property

Public Property Get auxIdCuentaCosto() As Long
    auxIdCuentaCosto = mProps.auxIdCuentaCosto
End Property

Public Property Get auxIdCuentaVenta() As Long
    auxIdCuentaVenta = mProps.auxIdCuentaVenta
End Property

Public Property Get auxIdCuentaDepreGasto() As Long
    auxIdCuentaDepreGasto = mProps.auxIdCuentaDepreGasto
End Property

Public Property Get auxIdCuentaDepreAcumulada() As Long
    auxIdCuentaDepreAcumulada = mProps.auxIdCuentaDepreAcumulada
End Property

Public Property Get auxIdCuentaRevaloriza() As Long
    auxIdCuentaRevaloriza = mProps.auxIdCuentaRevaloriza
End Property

Public Property Get auxIdCuentaDepRevaloriza() As Long
    auxIdCuentaDepRevaloriza = mProps.auxIdCuentaDepRevaloriza
End Property


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
        RecAfectaIvaItem = Me.GNComprobante.AFKardexRecAfectaIvaItem
        
        por = 0
        neto = Abs(Me.GNComprobante.AFKardexTotal(False))    'Total NETO sin recargo prorateado
        neto = neto - Me.GNComprobante.AFKardexDescItemTotal
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
'    If (value < 0) Or (value > 4) Then value = 0
    mProps.NumeroPrecio = value
End Property

'***Agregado. 02/sep/2003. Angel
Public Property Get NumeroPrecio() As Byte
    NumeroPrecio = mProps.NumeroPrecio
End Property

'***Agregado. 02/sep/2003. Angel
'***Para saber con que precio fue grabada la transacción
Private Sub BuscaNumeroPrecio()
    Dim iv As AFInventario
    
    Set iv = Me.GNComprobante.Empresa.RecuperaAFInventario(Me.IdInventario)
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
        Case Else
            Me.NumeroPrecio = 0
        End Select
    End If
    Set iv = Nothing
End Sub





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
    sql = "SELECT * FROM AFKardex WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !transid = mProps.objGNComprobante.transid
        !IdInventario = mProps.objGNComprobante.AFKardex(j).IdInventario
        !IdBodega = mProps.objGNComprobante.AFKardex(j).IdBodega
        !Cantidad = mProps.objGNComprobante.AFKardex(j).Cantidad
        !CostoTotal = mProps.objGNComprobante.AFKardex(j).CostoTotal
        !CostoRealTotal = mProps.objGNComprobante.AFKardex(j).CostoRealTotal
        !PrecioTotal = mProps.objGNComprobante.AFKardex(j).PrecioTotal
        !PrecioRealTotal = mProps.objGNComprobante.AFKardex(j).PrecioRealTotal
        !Descuento = mProps.objGNComprobante.AFKardex(j).Descuento
        !IVA = mProps.objGNComprobante.AFKardex(j).IVA
        !Orden = mProps.objGNComprobante.AFKardex(j).Orden
        !Nota = mProps.objGNComprobante.AFKardex(j).Nota
        !NumeroPrecio = mProps.objGNComprobante.AFKardex(j).NumeroPrecio
        !TiempoEntrega = mProps.objGNComprobante.AFKardex(j).TiempoEntrega
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

''AUC agregado para poder pasar cuando grabo la receta
'Friend Property Let IdInventario(ByVal value As Long)
'    mProps.IdInventario = value
'    mProps.objGNComprobante.SetModificado
'End Property



'*** jeaa 25/09/2006
Public Function CalcularSumaIvaItem( _
                    ByRef por As Double, _
                    ByRef bandCalculado As Boolean) As Currency
    Dim i As Long, RecAfectaIvaItem As Currency
    Dim neto As Currency, v As Currency
    
    If Not bandCalculado Then   'Para evitar que se repita muchas veces el mismo cálculo cuando llama del control IVRec
        'Sumar(restar) los recargos/descuentos que afecta a IVA por item
        RecAfectaIvaItem = Me.GNComprobante.AFKardexRecAfectaIvaItem
        
        por = 0
        neto = Abs(Me.GNComprobante.AFKardexTotal(False))    'Total NETO sin recargo prorateado
        neto = neto - Me.GNComprobante.AFKardexDescItemTotal
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

Public Property Let NumDepreReval(ByVal value As Currency)
    mProps.NumDepreReval = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get NumDepreReval() As Currency
    NumDepreReval = mProps.NumDepreReval
End Property

Public Property Let FechaIniDepReval(ByVal value As Date)
    mProps.FechaIniDepReval = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get FechaIniDepReval() As Date
    FechaIniDepReval = mProps.FechaIniDepReval
End Property

