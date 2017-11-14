VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "SiiTools  "
   ClientHeight    =   6555
   ClientLeft      =   135
   ClientTop       =   435
   ClientWidth     =   8535
   Icon            =   "Main.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stb1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   6285
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9393
            Key             =   "msg"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "29/08/2017"
            Key             =   "Fecha"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "16:31"
            Key             =   "Hora"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   1440
      Top             =   720
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuAbrirEmpresa 
         Caption         =   "Abrir &empresa..."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuConfigImpresora 
         Caption         =   "&Configurar impresora..."
      End
      Begin VB.Menu lin1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuHerramienta 
      Caption         =   "&Herramientas"
      Begin VB.Menu mnuCostoPri 
         Caption         =   "&Reprocesamiento de costos"
         Begin VB.Menu mnuCosto 
            Caption         =   "&Reprocesamiento de costos"
         End
         Begin VB.Menu mnuCostoxItem 
            Caption         =   "&Reprocesamiento de costos x Item"
         End
         Begin VB.Menu mnuRegeneraConsumos 
            Caption         =   "Regenera Consumos"
         End
         Begin VB.Menu mnuReasignaCostoIngreso 
            Caption         =   "Reasigna Costo Ingreso"
         End
      End
      Begin VB.Menu mnuAsientoPrin 
         Caption         =   "Re&generación de asientos"
         Begin VB.Menu mnuAsiento 
            Caption         =   "Re&generación de asientos"
         End
         Begin VB.Menu mnuAsientoxT 
            Caption         =   "Re&generación de asientos x Tran"
         End
         Begin VB.Menu mnuAsientoDuplicados 
            Caption         =   "Corrección Asientos Duplicados"
         End
         Begin VB.Menu mnuRegeneraAsientoRol 
            Caption         =   "Regeneracion  Asientos Roles"
         End
         Begin VB.Menu mnuGenerarUnAsiento 
            Caption         =   "Asientos x Lote"
         End
      End
      Begin VB.Menu mnuIsTeso 
         Caption         =   "Genera Igresos/Egresos Tesor."
         Begin VB.Menu mnuGenerarUnIngreso 
            Caption         =   "Generar Ingreso Automatico"
         End
         Begin VB.Menu mnuGenerarIngresoAuto 
            Caption         =   "Importar para Ingreso Automatico"
         End
         Begin VB.Menu mnuCruceTarjetas 
            Caption         =   "Genera cruce Tarjeta Cred."
         End
         Begin VB.Menu mnuCruceIVTarjetas 
            Caption         =   "Genera cruce IVTarjeta Cred."
         End
         Begin VB.Menu mnuGenerarPagos 
            Caption         =   "Generar Pagos Automatico"
         End
         Begin VB.Menu mnuVerificaPagos 
            Caption         =   "&Verifica Pagos Errados"
         End
         Begin VB.Menu mnuCambioCHP 
            Caption         =   "Cambio CHP"
         End
      End
      Begin VB.Menu mnuResumirIvK 
         Caption         =   "Re&sumir Kardex Inventario"
      End
      Begin VB.Menu mnuExist 
         Caption         =   "&Corrección de existencias"
      End
      Begin VB.Menu mnuCompr 
         Caption         =   "&Corrección de Comprometido"
      End
      Begin VB.Menu mnuNumSerie 
         Caption         =   "Correccion Exist NumSerie"
      End
      Begin VB.Menu mnuItemSinMovi 
         Caption         =   "&Eliminar items sin movimiento"
      End
      Begin VB.Menu mnuCentroSinMovi 
         Caption         =   "&Eliminar Centro sin movimiento"
      End
      Begin VB.Menu mnuTransErradas 
         Caption         =   "Buscar &Trans Erradas"
      End
      Begin VB.Menu mnuDesintegridad 
         Caption         =   "&Buscar Desintegridad"
      End
      Begin VB.Menu mnuIVFisico 
         Caption         =   "Constatación de Inventario"
         Begin VB.Menu mnuNuevo 
            Caption         =   "Nuevo"
         End
         Begin VB.Menu mnunada3m 
            Caption         =   "-"
         End
         Begin VB.Menu mnuConfiguracion 
            Caption         =   "Configuración"
         End
      End
      Begin VB.Menu mnuProduccion 
         Caption         =   "&Producción"
         Begin VB.Menu mnuReprocxPeriodo 
            Caption         =   "&Reproceso de Costos x Periodo"
         End
         Begin VB.Menu mnuCostoxProduccion 
            Caption         =   "&Reprocesamiento de Produccion"
         End
      End
      Begin VB.Menu mnuHLin0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTotalVentas 
         Caption         =   "Actualizar &Total de Ventas"
      End
      Begin VB.Menu mnuComprasProveedor 
         Caption         =   "Actualizar Compras x Proveedor"
      End
      Begin VB.Menu mnuItems 
         Caption         =   "&Inventarios"
         Begin VB.Menu mnuPrecio 
            Caption         =   "&Actualizar precios"
         End
         Begin VB.Menu mnuPrecioISO 
            Caption         =   "&Actualizar precios ISO"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuIVA 
            Caption         =   "Actualizar &IVA de ítems"
         End
         Begin VB.Menu mnuItemGrupos 
            Caption         =   "Asignar Grupos a Items "
         End
         Begin VB.Menu mnuCuentaItem 
            Caption         =   "Actualizar Cuentas contables de ítems"
         End
         Begin VB.Menu mnuItemFamilia 
            Caption         =   "Asignar Items a Familias"
         End
         Begin VB.Menu mnuItemFraccion 
            Caption         =   "Asignar Bandera para &Venta en Fracción"
         End
         Begin VB.Menu mnuItemVenta 
            Caption         =   "Asignar Bandera para &Venta "
         End
         Begin VB.Menu mnuCUI 
            Caption         =   "Asignar Costo Ultima Ingreso"
         End
         Begin VB.Menu mnuItemMIMMAX 
            Caption         =   "Asignar exist MIN/MAX"
         End
         Begin VB.Menu mnullenaIvexist 
            Caption         =   "Llena Tabla Existencias"
         End
         Begin VB.Menu mnuDiasRepo 
            Caption         =   "Asigna Dias Reposicion"
         End
         Begin VB.Menu mnuAjustes 
            Caption         =   "Ajustes Automaticos"
         End
         Begin VB.Menu mnuPorDescuento 
            Caption         =   "Asignar % Descuento"
         End
         Begin VB.Menu mnuPorComision 
            Caption         =   "Asignar % Comision"
         End
         Begin VB.Menu mnuExistenciaNegativa 
            Caption         =   "Encontrar Existencia Negativa"
         End
         Begin VB.Menu mnuCREF 
            Caption         =   "Asignar Costo Referencial"
         End
         Begin VB.Menu mnuArancel 
            Caption         =   "Asignar Arancel"
         End
         Begin VB.Menu mnuVerificaTransformacion 
            Caption         =   "Verifica Transformaciones"
         End
         Begin VB.Menu mnuVerificaNotaEntrega 
            Caption         =   "Verifica Notas Entrega"
         End
         Begin VB.Menu mnuFechaUltEgr 
            Caption         =   "Actualizar Fecha del Ultimo Egreso"
         End
         Begin VB.Menu mnuFechaUltIng 
            Caption         =   "Actualizar Fecha del Ultimo Ingreso"
         End
         Begin VB.Menu mnuVerificaITEMFactEelc 
            Caption         =   "Verificador de Datos Fact. elect"
         End
         Begin VB.Menu mnuCalculoBuffer 
            Caption         =   "Calculo Buffer"
         End
         Begin VB.Menu mnuCalculoBufferAlm 
            Caption         =   "Calculo Buffer Almacen"
         End
         Begin VB.Menu mnuClasificaItems 
            Caption         =   "Clasificador Items x Ventas"
         End
      End
      Begin VB.Menu mnuProv 
         Caption         =   "Proveedores"
         Begin VB.Menu mnuCuentaProveedor 
            Caption         =   "Actualizar Cuentas contables de Proveedores"
         End
         Begin VB.Menu mnuProvGrupos 
            Caption         =   "Asignar Grupos a Proveedores"
         End
         Begin VB.Menu mnuProvinProv 
            Caption         =   "Asignar Provincia a Proveedores"
         End
      End
      Begin VB.Menu mnuCli 
         Caption         =   "Clientes"
         Begin VB.Menu mnuCuentaCliente 
            Caption         =   "Actualizar Cuentas contables de Cliente"
         End
         Begin VB.Menu mnuCliGrupos 
            Caption         =   "Asignar Grupos a Clientes "
         End
         Begin VB.Menu mnuGarGrupos 
            Caption         =   "Asignar Grupos a Garantes"
         End
         Begin VB.Menu mnuGenHistorial 
            Caption         =   "Genera Historial"
         End
         Begin VB.Menu mnuTotalVentasProm 
            Caption         =   "Asignar Limite Crédito Prom/Vtas"
         End
         Begin VB.Menu mnuProvinCli 
            Caption         =   "Asignar Provincia"
         End
         Begin VB.Menu mnuVerificaCIRUC 
            Caption         =   "Verificador de CI/RUC"
         End
         Begin VB.Menu mnuAutocalificador 
            Caption         =   "AutoCalificador"
            Begin VB.Menu mnupcGrupoxMontoVenta 
               Caption         =   "Asignacionde de Grupo x Monto Ventas"
            End
            Begin VB.Menu mnupcGrupoxMontoVentaCobro 
               Caption         =   "Asignacion de Grupo x Monto Ventas y cobros"
            End
         End
         Begin VB.Menu mnuDINARDAP 
            Caption         =   "Asignacion Datos para DINARDAP"
         End
         Begin VB.Menu mnuVerificaCIRUCFactEelc 
            Caption         =   "Verificador de Datos Fact. elect"
         End
         Begin VB.Menu mnuVerificaemail 
            Caption         =   "Verificador de e-mail"
         End
         Begin VB.Menu mnuParroquias 
            Caption         =   "Verifica Parroquias"
         End
         Begin VB.Menu mnuContratos 
            Caption         =   "Verifica Contratos"
         End
         Begin VB.Menu mnasignaIVDescuento 
            Caption         =   "Asigna Descuentos x Item"
         End
         Begin VB.Menu mnuAsignaVendedor 
            Caption         =   "Asignacion Vendedor"
         End
         Begin VB.Menu mnucreacionAgencia 
            Caption         =   "Creacion de Agencias"
         End
      End
      Begin VB.Menu mnuEmpleados 
         Caption         =   "Empleados"
         Begin VB.Menu mnuAsginarCtaCtbEmp 
            Caption         =   "Asignar Cuentas Contables"
         End
         Begin VB.Menu mnuAsignaGruposEmp 
            Caption         =   "Asignar Grupos"
         End
         Begin VB.Menu mnuProvEmp 
            Caption         =   "Asignar Provincia"
         End
         Begin VB.Menu mnuDivideNombre 
            Caption         =   "Dividir Nombre"
         End
      End
      Begin VB.Menu mnuCtas 
         Caption         =   "&Contabilidad"
         Begin VB.Menu mnuCuentaLocal 
            Caption         =   "Actualizar Sucursales en Cuentas Contables "
         End
         Begin VB.Menu mnuCuentaPresup 
            Caption         =   "Actualizar Presupuesto en Cuentas Contables "
         End
         Begin VB.Menu mnuRelaCtaSC 
            Caption         =   "Relacionador Ctas SC"
         End
         Begin VB.Menu mnuRelaCtaFE 
            Caption         =   "Relacionador Ctas FE"
         End
         Begin VB.Menu mnuRelaCta101 
            Caption         =   "Relacionador Ctas 101"
         End
      End
      Begin VB.Menu mnuVend 
         Caption         =   "&Vendedores"
         Begin VB.Menu mnuactualComision 
            Caption         =   "Actualiza Comisiones x Vendedor"
         End
         Begin VB.Menu mnuactualComisionJefe 
            Caption         =   "Actualiza Comisiones x Jefe Ventas"
         End
         Begin VB.Menu mnuactualComisionItem 
            Caption         =   "Actualiza Comisiones x Vendedor x Item"
         End
         Begin VB.Menu mnuactualComisionGrupoItem 
            Caption         =   "Actualiza Comisiones x Vendedor x Grupo Item"
         End
      End
      Begin VB.Menu mnuCompro 
         Caption         =   "&Comprobantes"
         Begin VB.Menu mnuAutorizacionSRI 
            Caption         =   "Num. Autorizacion SRI"
         End
         Begin VB.Menu mnuVendedores 
            Caption         =   "Asignar Vendedores"
         End
         Begin VB.Menu mnuCreaTransAnulada 
            Caption         =   "Crea Trans Anuladas"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCorrexistDocum 
            Caption         =   "&Corrección de exist. Doc"
         End
         Begin VB.Menu mnuFormaPagoSRI 
            Caption         =   "Forma Pago SRI Compra"
         End
         Begin VB.Menu mnuRelacionCompro 
            Caption         =   "Relacion comprobantes"
         End
         Begin VB.Menu mnuAsigEmpleado 
            Caption         =   "Asignacion Empleado"
         End
         Begin VB.Menu mnuCambiaSecuencial_ 
            Caption         =   "Cambia Secuencia"
         End
         Begin VB.Menu mnuAsignarSerie 
            Caption         =   "Asignar Num Serie(Yolita,bellaluz)"
         End
         Begin VB.Menu mnuFormaCobroSRI 
            Caption         =   "Actualiza Forma Cobro SRI Vtas"
         End
         Begin VB.Menu mnuActualizaEsCopiaRTC 
            Caption         =   "Actualiza Original/Copia RTC"
         End
         Begin VB.Menu mnuCompraListaSRI 
            Caption         =   "Compara lista SRI"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCorrexistDocumRUTA 
            Caption         =   "Correccion exist Ruta"
         End
      End
      Begin VB.Menu mnuaf 
         Caption         =   "&Activos Fijos"
         Begin VB.Menu mnuCuentaAFItem 
            Caption         =   "Actualizar Cuentas contables de ítems"
         End
         Begin VB.Menu mnuAFItemGrupos 
            Caption         =   "Asignar Grupos a Items "
         End
         Begin VB.Menu mnuAFVidaUtil 
            Caption         =   "Asignar Vida Util"
         End
         Begin VB.Menu mnuGeneraDepre 
            Caption         =   "Generar Depreciación Mensual"
         End
         Begin VB.Menu mnuGeneraDepreReval 
            Caption         =   "Generar Depreciación - Mensual Reval"
         End
         Begin VB.Menu mnuAFExist 
            Caption         =   "&Corrección de No. Deprecia"
         End
         Begin VB.Menu mnullenaAFexist 
            Caption         =   "Llena Tabla Existencias Custodio"
         End
         Begin VB.Menu mnuAFExistC 
            Caption         =   "Corrección de existencias Custodios"
         End
         Begin VB.Menu mnuActualizaCustodioActivo 
            Caption         =   "Actualizar Custodio en Activo"
         End
         Begin VB.Menu mnuGeneraDepreANT 
            Caption         =   "DEP ANT"
         End
      End
      Begin VB.Menu mnuHLin2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImportar 
         Caption         =   "I&mportar datos"
         Begin VB.Menu mnuImportaPlan 
            Caption         =   "&Plan de cuenta"
         End
         Begin VB.Menu mnuImportaItem 
            Caption         =   "&Items"
         End
         Begin VB.Menu mnuImportaProv 
            Caption         =   "Pro&veedor"
         End
         Begin VB.Menu mnuImportaCli 
            Caption         =   "&Cliente"
         End
         Begin VB.Menu mnuEmpleado 
            Caption         =   "Empleados"
         End
         Begin VB.Menu mnuImportaAF 
            Caption         =   "&Activo Fijo"
         End
         Begin VB.Menu mnulineimp 
            Caption         =   "-"
         End
         Begin VB.Menu mnuImportaSaldoPV 
            Caption         =   "Cuentas por pagar"
         End
         Begin VB.Menu mnuImportaSaldoCL 
            Caption         =   "Cuentas por cobrar"
         End
         Begin VB.Menu mnuImportaSaldoEmp 
            Caption         =   "Cuentas por Pagar Emp"
         End
         Begin VB.Menu mnuImportarDiario 
            Caption         =   "A&siento Contable"
         End
         Begin VB.Menu mnuImportaInventaio 
            Caption         =   "Can&tidad de inventario"
         End
         Begin VB.Menu mnuImportaAFInventaio 
            Caption         =   "Can&tidad de Activos"
         End
         Begin VB.Menu mnuImportaAFInventaioC 
            Caption         =   "Can&tidad de Activos Custodio"
         End
         Begin VB.Menu mnuImportarNumSerie 
            Caption         =   "Cantidad NumSerie"
         End
         Begin VB.Menu mnulineaextra 
            Caption         =   "-"
         End
         Begin VB.Menu mnuImportaListaPrecios 
            Caption         =   "Actualizar Precios"
         End
         Begin VB.Menu mnuLineLocutorios 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnupcGrupo 
            Caption         =   "Grupo Prov/Cli"
            Begin VB.Menu mnupcgrupo1 
               Caption         =   "Grupo 1"
            End
            Begin VB.Menu mnupcgrupo2 
               Caption         =   "Grupo 2"
            End
            Begin VB.Menu mnupcgrupo3 
               Caption         =   "Grupo 3"
            End
            Begin VB.Menu mnupcgrupo4 
               Caption         =   "Grupo 4"
            End
         End
         Begin VB.Menu mnuivgrupo 
            Caption         =   "Grupo Items"
            Begin VB.Menu mnuivgrupo1 
               Caption         =   "Grupo 1"
            End
            Begin VB.Menu mnuivgrupo2 
               Caption         =   "Grupo 2"
            End
            Begin VB.Menu mnuivgrupo3 
               Caption         =   "Grupo 3"
            End
            Begin VB.Menu mnuivgrupo4 
               Caption         =   "Grupo 4"
            End
            Begin VB.Menu mnuivgrupo5 
               Caption         =   "Grupo 5"
            End
            Begin VB.Menu mnuivgrupo6 
               Caption         =   "Grupo 6"
            End
         End
         Begin VB.Menu mnuafgrupo 
            Caption         =   "Grupo Activo Fijo"
            Begin VB.Menu mnuafgrupo1 
               Caption         =   "Grupo 1"
            End
            Begin VB.Menu mnuafgrupo2 
               Caption         =   "Grupo 2"
            End
            Begin VB.Menu mnuafgrupo3 
               Caption         =   "Grupo 3"
            End
            Begin VB.Menu mnuafgrupo4 
               Caption         =   "Grupo 4"
            End
            Begin VB.Menu mnuafgrupo5 
               Caption         =   "Grupo 5"
            End
         End
         Begin VB.Menu MNUNADA55 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPlanPresup 
            Caption         =   "Plan Presupuesto"
         End
         Begin VB.Menu mnuImportarPRDiario 
            Caption         =   "Asiento Inicial Presup."
         End
         Begin VB.Menu mnuPlanSC 
            Caption         =   "Plan Cuentas SC"
         End
         Begin VB.Menu mnuPlanFE 
            Caption         =   "Plan Cuentas Flujo Efectivo"
         End
         Begin VB.Menu mnuplanEnferme 
            Caption         =   "Listado Enfermedades"
         End
      End
      Begin VB.Menu mnuHlin1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuimpresiones 
         Caption         =   "Impresiones"
         Begin VB.Menu mnuImpresionPorLote 
            Caption         =   "Im&presión por lote"
         End
         Begin VB.Menu mnuImpresionEstadoCuentas 
            Caption         =   "Impresion de Estado de Cuentas"
         End
         Begin VB.Menu mnuImpresionChequePorLote 
            Caption         =   "Impresión C&heque por lote"
         End
         Begin VB.Menu mnuImpresionEtiquetasPorLote 
            Caption         =   "Impresión E&tiqueta por Lote"
         End
         Begin VB.Menu mnuImpresionEtiquetasPorLoteProd 
            Caption         =   "Impresión E&tiqueta por Produccion"
         End
         Begin VB.Menu mnuImpresionEtiquetasPorLoteFact 
            Caption         =   "Impresión E&tiqueta por Facturacion"
         End
         Begin VB.Menu ImpRelDep 
            Caption         =   "Impresion Relacion Dependencia"
         End
         Begin VB.Menu mnuImpNoti 
            Caption         =   "Impresion Notificacion"
         End
         Begin VB.Menu mnuCambiaAcopiadorPorLote 
            Caption         =   "Cambio Acopiador"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuHlin11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCierres 
         Caption         =   "Cierres"
         Begin VB.Menu mnuCierre 
            Caption         =   "Cierre de ejercicio"
         End
         Begin VB.Menu mnuCierrePC 
            Caption         =   "Cierre de Periodo Contable "
         End
         Begin VB.Menu mnuCierreTransXEntregar 
            Caption         =   "Paso Trans. por Entregar "
         End
         Begin VB.Menu mnuCierreTransXEntregarF 
            Caption         =   "Paso Trans. por Entregar Familias"
         End
         Begin VB.Menu mnuCierreTransXEntregarI 
            Caption         =   "Paso Trans. por Entregar Items"
         End
         Begin VB.Menu mnuCierreTickxFact 
            Caption         =   "Paso Tickets x Fact"
         End
         Begin VB.Menu mnurevindice 
            Caption         =   "arregla indices"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPenProd 
            Caption         =   "Paso Trans. por Producir"
         End
      End
      Begin VB.Menu mnuSeparadorCostos 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPe 
         Caption         =   "Punto Equilibrio"
         Begin VB.Menu mnuCostos 
            Caption         =   "Base de Calculo Producción"
         End
         Begin VB.Menu mnuCostos1 
            Caption         =   "Base de Calculo Ventas"
         End
         Begin VB.Menu mnuCostos2016 
            Caption         =   "Base de Calculo Producción 2016"
         End
         Begin VB.Menu mnuCostos12016 
            Caption         =   "Base de Calculo Ventas 2016"
         End
      End
      Begin VB.Menu mnupeMP3 
         Caption         =   "Punto Equilibrio MP3"
         Begin VB.Menu mnuCostosVentasMP3 
            Caption         =   "Base de Calculo Ventas MP3"
         End
      End
      Begin VB.Menu mnulineafacxLote 
         Caption         =   "-"
      End
      Begin VB.Menu mnuflote 
         Caption         =   "Facturacion x Lote"
         Begin VB.Menu mnuLoteGrupo 
            Caption         =   "Lote Grupo"
         End
         Begin VB.Menu mnuLoteCli 
            Caption         =   "Lote Cliente"
         End
         Begin VB.Menu mnuNotificaciones 
            Caption         =   "Genera Notificaciones"
         End
         Begin VB.Menu mnuTransporte 
            Caption         =   "Transporte"
         End
         Begin VB.Menu mnuLoteCC 
            Caption         =   "Lote x CC"
         End
         Begin VB.Menu mnuDevol 
            Caption         =   "Devoluciones"
         End
         Begin VB.Menu mnuDevolxDscto 
            Caption         =   "DevolucionesxDscto"
         End
      End
      Begin VB.Menu mnuColas 
         Caption         =   "Colas Produccion"
         Begin VB.Menu mnuRegColas 
            Caption         =   "Regenerar Colas"
         End
      End
   End
   Begin VB.Menu mnuTransferir 
      Caption         =   "&Transferir"
      Begin VB.Menu mnuTransExportar 
         Caption         =   "&Exportar"
      End
      Begin VB.Menu mnuTransImportar 
         Caption         =   "&Importar"
      End
      Begin VB.Menu mnulin 
         Caption         =   "-"
      End
      Begin VB.Menu mnuconfigura 
         Caption         =   "&Configuracion"
      End
      Begin VB.Menu mnuLecturaMedidor 
         Caption         =   "Ingresar Lectura Medidor"
      End
   End
   Begin VB.Menu mnuDeclaraciones 
      Caption         =   "&Declaraciones"
      Begin VB.Menu mnuDeclaracionesI 
         Caption         =   "Declaraciones"
         Begin VB.Menu mnuAnexoTra2016 
            Caption         =   "Anexo Transaccional 2016"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuAnexoTra2016Consol 
            Caption         =   "Anexo Transaccional 2016 Consolidado"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuAnexoICE 
            Caption         =   "Anexo ICE"
         End
         Begin VB.Menu mnuAnexoRDEP 
            Caption         =   "Anexo RDEP"
         End
         Begin VB.Menu mnuFor101 
            Caption         =   "Formulario 101"
         End
         Begin VB.Menu mnuF1042010 
            Caption         =   "Formulario 104 "
         End
      End
      Begin VB.Menu mnuReporteDINARDAP 
         Caption         =   "DINARDAP"
      End
      Begin VB.Menu nada 
         Caption         =   "-"
      End
      Begin VB.Menu mnuposdata3M 
         Caption         =   "Posdata 3M"
         Visible         =   0   'False
         Begin VB.Menu mnuVentas3m 
            Caption         =   "Ventas"
         End
         Begin VB.Menu mnuExistencias3M 
            Caption         =   "Existencias"
         End
      End
      Begin VB.Menu mnuPuntoVentas 
         Caption         =   "Autorizaciones &Punto de Ventas"
      End
   End
   Begin VB.Menu mnuComElec 
      Caption         =   "Comprobantes &Electrónicos"
      Begin VB.Menu mnuGeneraXML 
         Caption         =   "Genera xml Electronico"
      End
      Begin VB.Menu mnuGeneraRide 
         Caption         =   "Genera RIDE pdf"
      End
      Begin VB.Menu mnurecuperaxmlBD 
         Caption         =   "Recupera XML de BD"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Ve&ntana"
      WindowList      =   -1  'True
      Begin VB.Menu mnuCerrarTodas 
         Caption         =   "&Cerrar todas"
      End
      Begin VB.Menu linV1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "Casca&da"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Mosaico &horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Mosaico &vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Organizar iconos"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Ay&uda"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&Acerca de ..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Sub DevLote_Click()
frmGenerarDevolucionxlote.Inicio
End Sub

Private Sub ImpRelDep_Click()
frmImprimePorLote.InicioRelDep
End Sub

Private Sub MDIForm_Load()
    Dim pos As Integer
    mnuHerramienta.Enabled = False
    mnuTransferir.Enabled = False
    
        'jeaa 14/03/2005
    'para version de impresion grafica
'    If RegGet(REGISTRO & ProductCode & "\ComprasProveedor") = "1" Then
        mnuComprasProveedor.Enabled = True
'    Else
'        mnuComprasProveedor.Enabled = False
'    End If
    
    'jeaa 14/03/2005
    'para version de impresion grafica
    
'''
'''
'''
'''    If RegGet(REGISTRO & ProductCode & "\PuntoEquilibrio") = "1" Then
'''        mnupe.Visible = True
'''    Else
'''        mnupe.Visible = False
'''    End If
'''
'''    If RegGet(REGISTRO & ProductCode & "\PuntoEquilibrioMP3") = "1" Then
'''        mnupeMP3.Visible = True
'''    Else
'''        mnupeMP3.Visible = False
'''    End If
'''
'''    If RegGet(REGISTRO & ProductCode & "\FormulariosSRI") = "1" Then
'''        mnuDeclaracionesI.Visible = True
'''    Else
'''        mnuDeclaracionesI.Visible = False
'''    End If
'''
    
End Sub

Private Sub mnu1032010_Click()
 Dim f As frmFormulario104
    Set f = BuscaForm("frmFormulario104", "F1032010")
    If f Is Nothing Then Set f = New frmFormulario104
    f.Inicio "F1032010"
End Sub

Private Sub mnasignaIVDescuento_Click()
'jeaa 21/07/2009
    Dim f As Form
    
    Set f = BuscaForm("frmTotalVentas", "")
    If f Is Nothing Then Set f = New frmTotalVentas
    f.InicioAsignaDescuentoxItem "DescxClixItem"
End Sub

Public Sub mnuAbrirEmpresa_Click()
    Dim BandModulo As Boolean, pos As Integer
    'Cierra todas las ventanas abiertas
    mnuCerrarTodas_Click
    mnuHerramienta.Enabled = False
    
    frmGNSelecEmpresa.Show vbModal, Me
    
    'Habilita el menu si está aberta una empresa
    mnuHerramienta.Enabled = Not (gobjMain.EmpresaActual Is Nothing)
    mnuTransferir.Enabled = Not (gobjMain.EmpresaActual Is Nothing)
    
    'Si no tiene permiso, deshabilita el menu correspondiente
    mnuCosto.Enabled = gobjMain.GrupoActual.PermisoActual.CatInventarioCostoVer
    mnuPrecio.Enabled = gobjMain.GrupoActual.PermisoActual.CatInventarioPrecioMod
    mnuItemSinMovi.Enabled = gobjMain.GrupoActual.PermisoActual.CatInventarioMod
    
    
    If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "MEGA") > 0 Then
        frmMain.mnuposdata3M.Visible = True
    End If
    
    pos = InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "HORMI")
    If pos > 0 Then
        frmMain.mnuCierreTransXEntregarF.Visible = True
    End If
   
    pos = InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "ISO")
    If pos > 0 Then
        frmMain.mnuCambiaAcopiadorPorLote.Visible = True
    End If

        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImportarXML")) > 0 Then
            If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImportarXML") = 1 Then
                frmMain.mnuCompraListaSRI.Visible = True
            Else
                frmMain.mnuCompraListaSRI.Visible = False
            End If
        Else
                frmMain.mnuCompraListaSRI.Visible = False
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
    
    
    mnuCREF.Enabled = gobjMain.UsuarioActual.BandSupervisor
    mnuCreaTransAnulada.Visible = gobjMain.UsuarioActual.BandSupervisor
    
    pos = InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "ISO")
    If pos > 0 Then
        frmMain.mnuPrecioISO.Visible = True
        mnuPrecio.Visible = False
    End If
    
    mnuVendedores.Visible = gobjMain.UsuarioActual.BandSupervisor

    
End Sub

Private Sub mnuactualComision_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmComisionesVendedor
    
    Set f = BuscaForm("frmComisionesVendedor", "COMI")
    If f Is Nothing Then Set f = New frmComisionesVendedor
    f.Inicio "COMI"
#End If
End Sub

Private Sub mnuactualComisionGrupoItem_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmComisionesVendedor
    Set f = BuscaForm("frmComisionesVendedor", "COMIXIVGITEM")
    If f Is Nothing Then Set f = New frmComisionesVendedor
    f.Inicio "COMIXIVGITEM"
#End If

End Sub

Private Sub mnuactualComisionJefe_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmComisionesVendedor
    
    Set f = BuscaForm("frmComisionesVendedor", "COMIJEFE")
    If f Is Nothing Then Set f = New frmComisionesVendedor
    f.Inicio "COMIJEFE"
#End If
End Sub

Private Sub mnuActualizaCustodioActivo_Click()
'En versión DAO no está soportada la función
'jeaa 04/03/05
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As Form
    
    Set f = BuscaForm("frmTotalVentas", "")
    If f Is Nothing Then Set f = New frmTotalVentas
    f.InicioCustodioxActivo "CustoxActivo"
#End If

End Sub

Private Sub mnuActualizaEsCopiaRTC_Click()

'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    Set f = BuscaForm("frmITEMIVA_Cuenta", "ESCOPIARTC")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "ESCOPIARTC"
#End If
End Sub

Private Sub mnuAFExist_Click()
    'frmCorrecExist.Inicio
    
    #If DAOLIB Then
        MsgBox MSGERR_NODAO, vbInformation
        Exit Sub
    #Else
        Dim f As Form
        
        Set f = BuscaForm("AFExistencias", "")
        If f Is Nothing Then Set f = New frmCorrecExist
        f.InicioDepreciaciones "AFExistencias"
    #End If

End Sub

Private Sub mnuAFExistC_Click()
    'frmCorrecExist.Inicio
    
    #If DAOLIB Then
        MsgBox MSGERR_NODAO, vbInformation
        Exit Sub
    #Else
        Dim f As Form
        
        Set f = BuscaForm("AFExistenciasCustodio", "")
        If f Is Nothing Then Set f = New frmCorrecExist
        f.Inicio "AFExistenciasCustodio"
    #End If
End Sub

Private Sub mnuAFItemGrupos_Click()
'En versión DAO no está soportada la función
'jeaa 24/09/04 asignacion de grupo a los items
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmAFItemIVA_Cuenta
    Set f = BuscaForm("frmITEMIVA_Cuenta", "ITEM_AFGRUPOS")
    If f Is Nothing Then Set f = New frmAFItemIVA_Cuenta
    f.Inicio "ITEM_AFGRUPOS"
#End If
End Sub

Private Sub mnuAFVidaUtil_Click()
'En versión DAO no está soportada la función
'jeaa 24/09/04 asignacion de grupo a los items
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmAFItemIVA_Cuenta
    Set f = BuscaForm("frmITEMIVA_Cuenta", "ITEM_VIDAUTIL")
    If f Is Nothing Then Set f = New frmAFItemIVA_Cuenta
    f.Inicio "ITEM_VIDAUTIL"
#End If
End Sub

Private Sub mnuAjustes_Click()
    'frmAjusteInventario.Inicio "AjusteAutomatico"
    frmInventarioFisicoNew.Inicio "AjusteAutomatico"
End Sub

Private Sub mnuAnexoICE_Click()
 Dim f As frmAnexoICE
    Set f = BuscaForm("frmAnexoICE", "ICE2016")
    If f Is Nothing Then Set f = New frmAnexoICE
    f.Inicio "ICE2016"
End Sub

Private Sub mnuAnexoRDEP_Click()
    Dim f As frmAnexoRDEP
    Set f = BuscaForm("frmAnexoRDEP", "RDEP")
    If f Is Nothing Then Set f = New frmAnexoRDEP
    f.Inicio "RDEP"
End Sub

Private Sub mnuAnexoT_Click()
    Dim f As frmAnexoTransaccional
    Set f = BuscaForm("frmAnexoTransaccional", "FAT")
    If f Is Nothing Then Set f = New frmAnexoTransaccional
    f.Inicio "FAT"
End Sub

Private Sub mnuAnexoTra2013_Click()
    Dim f As frmAnexoTransaccional2013
    Set f = BuscaForm("frmAnexoTransaccional2013", "FAT")
    If f Is Nothing Then Set f = New frmAnexoTransaccional2013
    f.Inicio "FAT"
End Sub

Private Sub mnuAnexoTra2015_Click()
 Dim f As frmAnexoTransaccional2015
    Set f = BuscaForm("frmAnexoTransaccional2015", "FAT2015")
    If f Is Nothing Then Set f = New frmAnexoTransaccional2015
    f.Inicio2015 "FAT2015"
End Sub

Private Sub mnuAnexoTra2016_Click()
 Dim f As frmAnexoTransaccional2016
    Set f = BuscaForm("frmAnexoTransaccional2016", "FAT2016")
    If f Is Nothing Then Set f = New frmAnexoTransaccional2016
    f.Inicio2016 "FAT2016"
End Sub

Private Sub mnuAnexoTra2016Consol_Click()
 Dim f As frmAnexoTransaccional2016Consol
    Set f = BuscaForm("frmAnexoTransaccional2016Consol", "FAT2016")
    If f Is Nothing Then Set f = New frmAnexoTransaccional2016Consol
    f.Inicio2016 "FAT2016"

End Sub

Private Sub mnuArancel_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "ARANCEL")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "ARANCEL"
#End If

End Sub

Private Sub mnuAsginarCtaCtbEmp_Click()
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "CUENTA_EMP")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "CUENTA_EMP"
#End If
End Sub

Private Sub mnuAsiento_Click()
    frmRegeneraAsiento.Inicio
End Sub

Private Sub mnuAsientoDuplicados_Click()
    frmRegeneraAsiento.InicioAsientoDuplicado
End Sub

Private Sub mnuAsientoxItem_Click()
    frmRegeneraAsientoxItem.Inicio
End Sub

Private Sub mnuAsientoxT_Click()
    frmRegeneraAsientoNew.Inicio
End Sub

Private Sub mnuAsientoPresup_Click()
    frmRegeneraAsientoPresup.Inicio
End Sub

Private Sub mnuAsigEmpleado_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    Set f = BuscaForm("frmITEMIVA_Cuenta", "EMPDOC") 'KARDEX DOCUMENTOS * ENTREGAR
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "EMPDOC"
#End If
End Sub

Private Sub mnuAsignaFechaIni_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "FECHAINICIAL")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "FECHAINICIAL"
#End If
End Sub

Private Sub mnuAsignaGruposEmp_Click()
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    Set f = BuscaForm("frmITEMIVA_Cuenta", "PCGRUPOS_EMP")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "PCGRUPOS_EMP"
#End If

End Sub

Private Sub mnuAsignarSerie_Click()
Dim f As frmGnSecuencia
    If f Is Nothing Then Set f = New frmGnSecuencia
    f.InicioSerie
End Sub

Private Sub mnuAsignaVendedor_Click()
'En versión DAO no está soportada la función
'jeaa 24/09/04 asignacion de grupo a los items
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    Set f = BuscaForm("frmITEMIVA_Cuenta", "PCCLI_VENDEDOR")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "PCCLI_VENDEDOR"
#End If

End Sub

Private Sub mnuAutorizacionSRI_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "SRI")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "SRI"
#End If
End Sub

Private Sub mnuCalculoBuffer_Click()
    Dim f As Form
    Set f = BuscaForm("frmTotalVentas", "")
    If f Is Nothing Then Set f = New frmTotalVentas
    If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "UTIL") > 0 Then
        f.InicioCalculoBufferUtilesa "CalculoBufferUtilesa"
    Else
        f.InicioCalculoBuffer "CalculoBuffer"
    End If
End Sub

Private Sub mnuCalculoBufferAlm_Click()
    Dim f As Form
    Set f = BuscaForm("frmTotalVentas", "")
    If f Is Nothing Then Set f = New frmTotalVentas
'    If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "UTIL") > 0 Then
'        f.InicioCalculoBufferUtilesa "CalculoBufferUtilesa"
'    Else
        f.InicioCalculoBufferxAlmacen "CalculoBufferxAlmacen"
'    End If
End Sub

Private Sub mnuCambiaSecuencial__Click()
Dim f As frmGnSecuencia
    If f Is Nothing Then Set f = New frmGnSecuencia
    f.Inicio

End Sub

Private Sub mnuCambioCHP_Click()
    Dim f As Form
    
    Set f = BuscaForm("CambioCHP", "CambioCHP")
    If f Is Nothing Then Set f = New frmRegeneraCHP
    f.Inicio

End Sub

Private Sub mnuCentroSinMovi_Click()
    frmItemSinMovi.InicioCentroCosto
End Sub

Private Sub mnuCierre_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#End If
    
    frmCierre.Inicio
End Sub

Private Sub mnuCierrePC_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#End If
    
    frmCierrePeriodo.Inicio
End Sub

Private Sub mnuCierreTickxFact_Click()
'En versión DAO no está soportada la función
'jeaa 06/06/2007
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#End If
    frmCierreTransXTicket.Inicio ("Ingreso")
End Sub

Private Sub mnuCierreTransXEntregar_Click()
'En versión DAO no está soportada la función
'jeaa 06/06/2007
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#End If
    frmCierreTransXEntregar.Inicio ("Items")
End Sub

Private Sub mnuCierreTransXEntregarF_Click()
'En versión DAO no está soportada la función
'jeaa 06/06/2007
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#End If
    frmCierreTransXEntregar.Inicio ("Familias")
End Sub

Private Sub mnuCierreTransXEntregarI_Click()
'En versión DAO no está soportada la función
'jeaa 06/06/2007
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#End If
    frmCierreTransXEntregar.Inicio ("ItemsHormi")
End Sub

Private Sub mnuClasificaItems_Click()
'En versión DAO no está soportada la función
'jeaa 04/03/05
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As Form
    
    Set f = BuscaForm("frmTotalVentas", "")
    If f Is Nothing Then Set f = New frmTotalVentas
    f.InicioVentasxItemxSuc "VxIxSuc"
#End If

End Sub

Private Sub mnuCliGrupos_Click()
'En versión DAO no está soportada la función
'jeaa 24/09/04 asignacion de grupo a los items
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    Set f = BuscaForm("frmITEMIVA_Cuenta", "PCGRUPOS_CLI")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "PCGRUPOS_CLI"
#End If
End Sub

Private Sub mnuCompr_Click()
'    frmCorrecExist.InicioComprometido "Comprometido"
    
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As Form
    
    Set f = BuscaForm("Comprometido", "")
    If f Is Nothing Then Set f = New frmCorrecExist
    f.InicioComprometido "Comprometido"
#End If
    
    
End Sub

Private Sub mnuCompraListaSRI_Click()

'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmListaSRI
    
    Set f = BuscaForm("frmListaSRI", "Compras")
    If f Is Nothing Then Set f = New frmListaSRI
    f.Inicio "Compras"
    
#End If

End Sub

Private Sub mnuComprasProveedor_Click()
'En versión DAO no está soportada la función
'jeaa 04/03/05
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As Form
    
    Set f = BuscaForm("frmTotalVentas", "")
    If f Is Nothing Then Set f = New frmTotalVentas
    f.InicioComprasProveedor "CxTrans"
#End If
End Sub

Private Sub mnuComprometido_Click()
    frmGenerarCompAutomatico.Inicio
End Sub

Private Sub mnuConfigImpresora_Click()
    On Error Resume Next
    With dlg1
        .DialogTitle = "Configurar página"
        .CancelError = True
        .ShowPrinter
    End With
End Sub

Public Sub mnuCerrarTodas_Click()
    Dim frm As Form
    
    'Descarga todas las ventanas excepto frmMain misma.
    For Each frm In Forms
        If Not (frm Is Me) Then Unload frm
    Next frm
    Set frm = Nothing
End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Asegura que no se quede ningún formulario en la memoria
    mnuCerrarTodas_Click
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    gobjMain.EmpresaActual.Cerrar
    gobjMain.Cerrar
    Set gobjMain = Nothing
End Sub

Public Sub CambiaCaption()
    'Muestra nombre de empresa y nombre de operador
    Dim pos As Integer
    Me.Caption = App.Title & " "
    If Not gobjMain.EmpresaActual Is Nothing Then
        Me.Caption = Me.Caption & "[" & gobjMain.EmpresaActual.Descripcion & "]"
    End If
    If Not gobjMain.UsuarioActual Is Nothing Then
        Me.Caption = Me.Caption & " (" & gobjMain.UsuarioActual.NombreUsuario & ")"
    End If
    
    If gobjMain.ModoDemo Then
        Me.Caption = Me.Caption & " ** DEMO **"
    End If
    
    pos = InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "HORMI")
    If pos > 0 Then
        frmMain.mnuCierreTransXEntregarF.Visible = True
        frmMain.mnuCierreTransXEntregarI.Visible = True
    Else
        frmMain.mnuCierreTransXEntregarF.Visible = False
        frmMain.mnuCierreTransXEntregarI.Visible = False
    End If
    
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImportarXML")) > 0 Then
        If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImportarXML") = 1 Then
            frmMain.mnuCompraListaSRI.Visible = True
        Else
            frmMain.mnuCompraListaSRI.Visible = False
        End If
    Else
            frmMain.mnuCompraListaSRI.Visible = False
    End If

    

''    pos = InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "AERO")
''    If pos > 0 Then
''        frmMain.mnuaereo.Visible = True
''    Else
''        frmMain.mnuaereo.Visible = False
''    End If
    
    pos = InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "ISO")
    If pos > 0 Then
        frmMain.mnuPrecioISO.Visible = True
        mnuPrecio.Visible = False
        frmMain.mnuCambiaAcopiadorPorLote.Visible = True
    End If

    If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "MEGA") > 0 Then
        frmMain.mnuposdata3M.Visible = True
    End If

    
End Sub

Private Sub mnuconfigura_Click()
    frmConfig.Inicio
End Sub

Private Sub mnuConfiguracion_Click()
    frmConfigIVFisico.Inicio
End Sub

Private Sub mnuContratos_Click()
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As Form
    
    Set f = BuscaForm("Contratos", "")
    If f Is Nothing Then Set f = New frmRegeneraTransformacion
    f.InicioContratos "Contratos"
#End If

End Sub

Private Sub mnuCorrexistDocum_Click()
    'frmCorrecExist.Inicio
    
    #If DAOLIB Then
        MsgBox MSGERR_NODAO, vbInformation
        Exit Sub
    #Else
        Dim f As Form
        
        Set f = BuscaForm("ExistenciasDocum", "")
        If f Is Nothing Then Set f = New frmCorrecExist
        f.InicioDocum "ExistenciasDocum"
    #End If
    

End Sub

Private Sub mnuCorrexistDocumRUTA_Click()
'frmCorrecExist.Inicio
    #If DAOLIB Then
        MsgBox MSGERR_NODAO, vbInformation
        Exit Sub
    #Else
        Dim f As Form
        
        Set f = BuscaForm("ExistenciasDocumRuta", "")
        If f Is Nothing Then Set f = New frmCorrecExist
        f.InicioDocum "ExistenciasDocumRuta"
    #End If
End Sub

Private Sub mnuCosto_Click()
    frmReprocCosto.Inicio
End Sub

Private Sub mnuCostoNew_Click()
    frmReprocCostoQuick.Inicio
End Sub

Private Sub mnuCostos_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemCostos
    
    Set f = BuscaForm("frmItemCostos", "Costos")
    If f Is Nothing Then Set f = New frmItemCostos
    f.Inicio "Costos"
    
#End If

End Sub

Private Sub mnuCostos1_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemCostos
    
    Set f = BuscaForm("frmItemCostos", "Produccion")
    If f Is Nothing Then Set f = New frmItemCostos
    f.Inicio "Produccion"
    
#End If
End Sub

Private Sub mnuCostos12016_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemCostos2016
    
    Set f = BuscaForm("frmItemCostos2016", "Produccion")
    If f Is Nothing Then Set f = New frmItemCostos2016
    f.Inicio "Produccion"
    
#End If
End Sub

Private Sub mnuCostos2016_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemCostos2016
    
    Set f = BuscaForm("frmItemCostos2016", "Costos")
    If f Is Nothing Then Set f = New frmItemCostos2016
    f.Inicio "Costos"
    
#End If
End Sub

Private Sub mnuCostosVentasMP3_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmPuntoEquilibrioMP3
    
    Set f = BuscaForm("frmPuntoEquilibrioMP3", "Ventas")
    If f Is Nothing Then Set f = New frmPuntoEquilibrioMP3
    f.Inicio "Ventas"
    
#End If
End Sub

Private Sub mnuCostoxItem_Click()
    frmReprocCostoxItem.Inicio
End Sub

Private Sub mnuCostoxProduccion_Click()
    frmReprocCostoxProduccion.Inicio
End Sub

Private Sub mnucreacionAgencia_Click()
'En versión DAO no está soportada la función
'jeaa 13/04/2005
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "PCAGENCIA")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "PCAGENCIA"
#End If

End Sub

Private Sub mnuCreaTransAnulada_Click()
    frmGenerarCompAutomatico.InicioAnulados ("InicioAnulados")
End Sub

Private Sub mnuCREF_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "COSTOREF")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "COSTOREF"
#End If
End Sub

Private Sub mnuCruceIVTarjetas_Click()
    frmGenerarIngresoAutomaticoTC.InicioCruceIVTarjetas "CruceIVTarjetas"
End Sub

Private Sub mnuCruceTarjetas_Click()
    frmGenerarIngresoAutomaticoTC.InicioCruceTarjetas "CruceTarjetas"
End Sub

Private Sub mnuCuentaAFItem_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmAFItemIVA_Cuenta
    
    Set f = BuscaForm("frmAFITEMIVA_Cuenta", "CUENTA")
    If f Is Nothing Then Set f = New frmAFItemIVA_Cuenta
    f.Inicio "CUENTA"
#End If
End Sub

Private Sub mnuCuentaCliente_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "CUENTA_CLI")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "CUENTA_CLI"
#End If
End Sub

Private Sub mnuCuentaItem_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "CUENTA")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "CUENTA"
#End If
End Sub

Private Sub mnuCuentaLocal_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "CUENTA_LOCAL")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "CUENTA_LOCAL"
#End If
End Sub

Private Sub mnuCuentaPresup_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "CUENTA_PRESUP")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "CUENTA_PRESUP"
#End If
End Sub

Private Sub mnuCuentaProveedor_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "CUENTA_PROV")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "CUENTA_PROV"
#End If
End Sub

'jeaa 01/3/2007
Private Sub mnuCUI_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "COSTOUI")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "COSTOUI"
#End If
End Sub

Private Sub mnuDesintegridad_Click()
    frmB_Desintegridad.Inicio
End Sub

Private Sub mnuDevol_Click()
frmFacturarxLote.InicioDevolucion
End Sub

Private Sub mnuDevolxDscto_Click()
frmFacturarxLote.InicioDevolucionxDscto
End Sub

Private Sub mnuDiasRepo_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "DIASREPO")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "DIASREPO"
#End If

End Sub

Private Sub mnuDINARDAP_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "DINARDAP")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "DINARDAP"
#End If
End Sub

Private Sub mnuDivideNombre_Click()
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    Set f = BuscaForm("frmITEMIVA_Cuenta", "DIVNOMEMP")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "DIVNOMEMP"
#End If

End Sub

Private Sub mnuExist_Click()
    'frmCorrecExist.Inicio
    
    #If DAOLIB Then
        MsgBox MSGERR_NODAO, vbInformation
        Exit Sub
    #Else
        Dim f As Form
        
        Set f = BuscaForm("Existencias", "")
        If f Is Nothing Then Set f = New frmCorrecExist
        f.Inicio "Existencias"
    #End If
    
    
End Sub

Private Sub mnuExistenciaNegativa_Click()
'En versión DAO no está soportada la función
'jeaa 13/04/2005
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "IVEXISTNEG")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "IVEXISTNEG"
#End If
End Sub

Private Sub mnuExistencias3M_Click()
    Dim f As frm3M
    Set f = BuscaForm("frm3M", "")
    If f Is Nothing Then Set f = New frm3M
    f.InicioExist ""
End Sub

Private Sub mnuF103_Click()
    Dim f As frmFormulario104
    Set f = BuscaForm("frmFormulario104", "F104")
    If f Is Nothing Then Set f = New frmFormulario104
    f.Inicio "F103"
End Sub

Private Sub mnuF104_Click()
    Dim f As frmFormulario104
    Set f = BuscaForm("frmFormulario104", "F104")
    If f Is Nothing Then Set f = New frmFormulario104
    f.Inicio "F104"
End Sub

Private Sub mnuF1042010_Click()
    Dim f As frmFormulario1042014
   Set f = BuscaForm("frmFormulario104", "F1042010")
    If f Is Nothing Then Set f = New frmFormulario1042014
   f.Inicio "F1042010"
End Sub

Private Sub mnuFactAerolinieas_Click()
    'jeaa 21/01/2009
    frmGenerarFacturaAeropuerto.Inicio
End Sub

Private Sub mnuFacturarxLote_Click()
    frmFacturarxLote.Inicio
End Sub

Private Sub mnuFechaUltEgr_Click()
    Dim f As Form
    Set f = BuscaForm("frmTotalVentas", "")
    If f Is Nothing Then Set f = New frmTotalVentas
    f.InicioFechaUltimoEgreso "FechaUltimoEgreso"
End Sub

Private Sub mnuFechaUltIng_Click()
    Dim f As Form
    Set f = BuscaForm("frmTotalVentas", "")
    If f Is Nothing Then Set f = New frmTotalVentas
    f.InicioFechaUltimoIngreso "FechaUltimoIngreso"

End Sub

Private Sub mnuFor101_Click()
    Dim f As frmReporte101
    
    Set f = BuscaForm("frmReporte101", "F101")
    If f Is Nothing Then Set f = New frmReporte101
    f.Inicio "F101"

End Sub

Private Sub mnuFormaCobroSRI_Click()

'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    Set f = BuscaForm("frmITEMIVA_Cuenta", "FORMASRI")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "FORMASRI"
#End If
End Sub

Private Sub mnuFormaPagoSRI_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "FORMAPAGOSRI")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "FORMAPAGOSRI"
#End If

End Sub

Private Sub mnuGenAlquilados_Click()
frmGeneraAlquilados.Inicio
End Sub

Private Sub mnuGarGrupos_Click()
'En versión DAO no está soportada la función
'jeaa 24/09/04 asignacion de grupo a los items
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    Set f = BuscaForm("frmITEMIVA_Cuenta", "PCGRUPOS_GAR")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "PCGRUPOS_GAR"
#End If
End Sub

Private Sub mnuGeneraDepre_Click()
    frmGenerarDepreciacion.Inicio "Depre"
End Sub

Private Sub mnuGeneraDepreANT_Click()
    frmGenerarDepreciacion.Inicio "DepreANT"
End Sub

Private Sub mnuGeneraDepreReval_Click()
    frmGenerarDepreciacion.InicioReval "DepreReval"
End Sub

Private Sub mnuGeneraDepreRevalA_Click()
    frmGenerarDepreciacion.InicioReval "DepreRevalA"
End Sub

Private Sub mnuGeneraRide_Click()
    frmImprimePorLote.InicioPDF
End Sub

Private Sub mnuGenerarIngresoAuto_Click()
    frmGenerarIngresoAutomaticoTC.InicioImportar "JEP"
End Sub

Private Sub mnuGenerarPagos_Click()
    frmGenerarIngresoAutomaticoTC.InicioPago "Pago"
End Sub

Private Sub mnuGenerarUnAsiento_Click()
    frmGenerarUnAsientoxLoteDeTransacciones.Inicio
End Sub

Private Sub mnuGenerarUnAsientoPresup_Click()
''    frmGenerarUnAsientoxLoteDeTransaccionesPresup.Inicio
End Sub

Private Sub mnuGenerarUnIngreso_Click()
    frmGenerarIngresoAutomaticoTC.Inicio
End Sub

Private Sub mnuGeneraXML_Click()
    frmImprimePorLote.InicioXML
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer


    'si no hay archivo de ayuda para este proyecto, mostrar un mensaje al usuario
    'puede establecer el archivo de Ayuda para su aplicación en el cuadro
    'de diálogo Propiedades del proyecto
    If Len(App.HelpFile) = 0 Then
        MsgBox "No se puede mostrar el contenido de la Ayuda. No hay Ayuda asociada a este proyecto.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer


    'si no hay archivo de ayuda para este proyecto, mostrar un mensaje al usuario
    'puede establecer el archivo de Ayuda para la aplicación en el cuadro
    'de diálogo Propiedades del proyecto
    If Len(App.HelpFile) = 0 Then
        MsgBox "No se puede mostrar el contenido de la Ayuda. No hay Ayuda asociada a este proyecto.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub



Private Sub mnuImpNoti_Click()
 frmGeneraNotificacion.InicioImpresion
End Sub

Private Sub mnuImportaAF_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmImportacion
    
    Set f = BuscaForm("frmImportacion", "AFITEM")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "AFITEM"
    
#End If
End Sub

Private Sub mnuImportaAFInventaio_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmImportacion
    
    Set f = BuscaForm("frmImportacion", "AFINVENTARIO")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "AFINVENTARIO"
    
#End If
End Sub

Private Sub mnuImportaAFInventaioC_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmImportacion
    
    Set f = BuscaForm("frmImportacion", "AFINVENTARIOC")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "AFINVENTARIOC"
    
#End If
End Sub

Private Sub mnuImportaCli_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmImportacion

    Set f = BuscaForm("frmImportacion", "PCCLI")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "PCCLI"
    
#End If
End Sub


Private Sub mnuImportaInventaio_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmImportacion
    
    Set f = BuscaForm("frmImportacion", "INVENTARIO")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "INVENTARIO"
    
#End If
    
End Sub



Private Sub mnuImportaItem_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmImportacion
    
    Set f = BuscaForm("frmImportacion", "ITEM")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "ITEM"
    
#End If
End Sub


Private Sub mnuImportaListaPrecios_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmImportacion
    
    Set f = BuscaForm("frmImportacion", "EXISTENCIA MINIMA")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "EXISTENCIA MINIMA"
    
#End If

End Sub

Private Sub mnuImportaPlan_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmImportacion
    
    Set f = BuscaForm("frmImportacion", "PLANCUENTA")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "PLANCUENTA"
    
#End If
End Sub

Private Sub mnuImportaProv_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmImportacion
    
    Set f = BuscaForm("frmImportacion", "PCPROV")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "PCPROV"
    
#End If
End Sub

Private Sub mnuImportarDiario_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmImportacion
    
    Set f = BuscaForm("frmImportacion", "DIARIO")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "DIARIO"
    
#End If
End Sub

Private Sub mnuImportarPRDiario_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmImportacion
    
    Set f = BuscaForm("frmImportacion", "PRDIARIO")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "PRDIARIO"
    
#End If
End Sub

Private Sub mnuImportaSaldoCEmp_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmImportacion
    Set f = BuscaForm("frmImportacion", "PORPAGAREMP")
   If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "PORCOBRAREMP"
#End If
End Sub

Private Sub mnuImportaSaldoCL_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmImportacion
    
    Set f = BuscaForm("frmImportacion", "PORCOBRAR")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "PORCOBRAR"
    
#End If
End Sub

Private Sub mnuImportaSaldoEmp_Click()
    #If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmImportacion
    Set f = BuscaForm("frmImportacion", "PORPAGAREMP")
   If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "PORPAGAREMP"
#End If
End Sub

Private Sub mnuImportaSaldoPV_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmImportacion
    
    Set f = BuscaForm("frmImportacion", "PORPAGAR")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "PORPAGAR"
    
#End If
End Sub


Private Sub mnuImpresionChequePorLote_Click()
    'jeaa 12/11/2008
    frmImprimePorLote.InicioCheque
End Sub

Private Sub mnuImpresionEstadoCuentas_Click()
    frmImprimePorLoteEstadoCuenta.Inicio
End Sub

Private Sub mnuImpresionEtiquetasPorLote_Click()
    'jeaa 01/10/2009
    frmImprimePorLote.InicioEtiqueta
End Sub

Private Sub mnuImpresionEtiquetasPorLoteFact_Click()
    frmImprimePorLote.InicioEtiquetaFacturacion
End Sub

Private Sub mnuImpresionEtiquetasPorLoteProd_Click()
    frmImprimePorLote.InicioEtiquetaProduccion
End Sub

Private Sub mnuImpresionPorLote_Click()
    frmImprimePorLote.Inicio
End Sub

Private Sub mnuItemArea_Click()
'En versión DAO no está soportada la función
'jeaa 15/09/2005
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "AREA")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "AREA"
#End If
End Sub

Private Sub mnuItemFamilia_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    Set f = BuscaForm("frmITEMIVA_Cuenta", "ITEM_FAMILIA")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "ITEM_FAMILIA"
#End If
End Sub

Private Sub mnuItemFraccion_Click()
'En versión DAO no está soportada la función
'jeaa 13/04/2005
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "FRACCION")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "FRACCION"
#End If
End Sub

Private Sub mnuItemGrupos_Click()
'En versión DAO no está soportada la función
'jeaa 24/09/04 asignacion de grupo a los items
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    Set f = BuscaForm("frmITEMIVA_Cuenta", "ITEM_IVGRUPOS")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "ITEM_IVGRUPOS"
#End If
End Sub

Private Sub mnuItemMIMMAX_Click()
'En versión DAO no está soportada la función
'jeaa 16/02/2008
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "MINMAX")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "MINMAX"
#End If
End Sub

Private Sub mnuItemSinMovi_Click()
    frmItemSinMovi.Inicio
End Sub

Private Sub mnuItemVenta_Click()
'En versión DAO no está soportada la función
'jeaa 26/12/2005
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "VENTA")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "VENTA"
#End If

End Sub

Private Sub mnuIVA_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "IVA")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "IVA"
#End If
End Sub

Private Sub mnuivgrupo6_Click()
    Dim f As frmImportacion
    Set f = BuscaForm("frmImportacion", "IVGRUPO6")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "IVGRUPO6"
End Sub

Private Sub mnuLecturaMedidor_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "LECTURAS")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "LECTURAS"
#End If

End Sub

Private Sub mnuLiquidar_Click()
    frmLiquidar.Inicio
End Sub

Private Sub mnullenaAFexist_Click()
'En versión DAO no está soportada la función
'jeaa 13/04/2005
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "AFEXIST")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "AFEXIST"
#End If
End Sub

Private Sub mnullenaIvexist_Click()
'En versión DAO no está soportada la función
'jeaa 13/04/2005
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "IVEXIST")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "IVEXIST"
#End If
End Sub


Private Sub mnuLoteCC_Click()
    frmFacturarxLote.InicioxCC
End Sub

Private Sub mnuLoteCli_Click()
    frmFacturarxLote.InicioxCliente
End Sub

Private Sub mnuLoteGrupo_Click()
    frmFacturarxLote.Inicio
End Sub


Private Sub mnuNotificaciones_Click()
    frmGeneraNotificacion.Inicio
End Sub

Private Sub mnuNuevo_Click()
    frmInventarioFisico.Inicio "InventarioFisico"
End Sub

Private Sub mnuNumSerie_Click()
    #If DAOLIB Then
        MsgBox MSGERR_NODAO, vbInformation
        Exit Sub
    #Else
        Dim f As Form
        Set f = BuscaForm("ExistenciasSerie", "")
        If f Is Nothing Then Set f = New frmCorrecExist
        f.Inicio "ExistenciasSerie"
    #End If
End Sub

Private Sub mnuParroquias_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "PCPARR")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "PCPARR"
#End If
End Sub

Private Sub mnupcgrupo1_Click()
    Dim f As frmImportacion
    Set f = BuscaForm("frmImportacion", "PCGRUPO1")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "PCGRUPO1"
End Sub

Private Sub mnupcgrupo2_Click()
    Dim f As frmImportacion
    Set f = BuscaForm("frmImportacion", "PCGRUPO2")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "PCGRUPO2"

End Sub

Private Sub mnupcgrupo3_Click()
    Dim f As frmImportacion
    Set f = BuscaForm("frmImportacion", "PCGRUPO3")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "PCGRUPO3"
End Sub


Private Sub mnupcgrupo4_Click()
    Dim f As frmImportacion
    Set f = BuscaForm("frmImportacion", "PCGRUPO4")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "PCGRUPO4"
End Sub

Private Sub mnupcGrupoxMontoVenta_Click()
'jeaa 21/07/2009
    Dim f As Form
    
    Set f = BuscaForm("frmTotalVentas", "")
    If f Is Nothing Then Set f = New frmTotalVentas
    f.InicioPcGrupoxMontoVentas "PCGxMontoVenta"
End Sub

Private Sub mnupcGrupoxMontoVentaCobro_Click()
    'AUC   2017
    Dim f As Form
    Set f = BuscaForm("frmTotalVentas", "")
    If f Is Nothing Then Set f = New frmTotalVentas
    f.InicioPcGrupoxMontoVentasCobros "PCGxMontoVentaCobro"
End Sub

Private Sub mnuPenProd_Click()
'AUC 22/05/08 para transferir alquileres que no se develven todavia a otra empresa
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#End If
    frmCierreTransXDevolver.InicioProduccion
End Sub

Private Sub mnuplanEnferme_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmImportacion
    
    Set f = BuscaForm("frmImportacion", "PLANENFERME")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "PLANENFERME"
    
#End If
End Sub

Private Sub mnuPlanFE_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmImportacion
    
    Set f = BuscaForm("frmImportacion", "PLANCUENTAFE")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "PLANCUENTAFE"
    
#End If
End Sub

Private Sub mnuPlanPresup_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmImportacion
    
    Set f = BuscaForm("frmImportacion", "PLANCUENTA")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "PLANPRCUENTA"
    
#End If
End Sub

Private Sub mnuPlanSC_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmImportacion
    
    Set f = BuscaForm("frmImportacion", "PLANCUENTASC")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "PLANCUENTASC"
    
#End If

End Sub

Private Sub mnuPorComision_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "PORCOMI")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "PORCOMI"
#End If

End Sub

Private Sub mnuPorDescuento_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "PORDESC")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "PORDESC"
#End If

End Sub

Private Sub mnuPrecio_Click()
    frmPrecios.Inicio
End Sub

Private Sub mnuPrecioISO_Click()
    frmPreciosISO.Inicio
End Sub

Private Sub mnuProvEmp_Click()
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    Set f = BuscaForm("frmITEMIVA_Cuenta", "PROVINCIASEMP")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "PROVINCIASEMP"
#End If
End Sub

Private Sub mnuProvGrupos_Click()
'En versión DAO no está soportada la función
'jeaa 24/09/04 asignacion de grupo a los items
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    Set f = BuscaForm("frmITEMIVA_Cuenta", "PCGRUPOS_PROV")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "PCGRUPOS_PROV"
#End If
End Sub

Private Sub mnuProvinCli_Click()
'En versión DAO no está soportada la función
'jeaa 24/09/04 asignacion de grupo a los items
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    Set f = BuscaForm("frmITEMIVA_Cuenta", "PROVINCIAS")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "PROVINCIAS"
#End If
End Sub

Private Sub mnuProvinProv_Click()
'En versión DAO no está soportada la función
'jeaa 24/09/04 asignacion de grupo a los items
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    Set f = BuscaForm("frmITEMIVA_Cuenta", "PROVINCIAS_PROV")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "PROVINCIAS_PROV"
#End If
End Sub

Private Sub mnuPuntoVentas_Click()
    frmIVGenAutoSRI.Inicio
End Sub

Private Sub mnuReasignaCostoIngreso_Click()
    frmReprocCosto.InicioCostoxProveedor "CostoxProveedor"
End Sub

Private Sub mnuRecetaxItem_Click()
    frmRegeneraRecetasVentas.Inicio
End Sub

Private Sub mnurecuperaxmlBD_Click()
    frmImprimePorLote.InicioRecuperaXMLBD
End Sub

Private Sub mnuRegColas_Click()
frmGeneraAlquilados.InicioColas ("Colas")

End Sub

Private Sub mnuRegeneraAsientoRol_Click()
    frmRegeneraAsientoNew.InicioRoles
End Sub

Private Sub mnuRegeneraConsumos_Click()
    frmRegeneraConsumos.Inicio
End Sub

Private Sub mnuRegeneraDesperdicio_Click()
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As Form
    
    Set f = BuscaForm("Desperdicio", "")
    If f Is Nothing Then Set f = New frmRegeneraDesperdicio
    f.InicioDesperdicio "Desperdicio"
#End If

End Sub

Private Sub mnuRelacionCompro_Click()

'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    Set f = BuscaForm("frmITEMIVA_Cuenta", "COMPROB")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "COMPROB"
#End If
End Sub

Private Sub mnuRelaCta101_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "CUENTA101")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "CUENTA101"
#End If
End Sub

Private Sub mnuRelaCtaFE_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "CUENTAFE")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "CUENTAFE"
#End If
End Sub

Private Sub mnuRelaCtaSC_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    
    Set f = BuscaForm("frmITEMIVA_Cuenta", "CUENTASC")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "CUENTASC"
#End If
End Sub

Private Sub mnuReporteDINARDAP_Click()
    Dim f As frmDinardap
    Set f = BuscaForm("frmDinardap", "")
    If f Is Nothing Then Set f = New frmDinardap
    f.Inicio ""
End Sub

Private Sub mnuReprocxPeriodo_Click()
    frmReprocCostoPeriodo.Inicio
End Sub

Private Sub mnuResumirIvK_Click()
    frmResumir.InicioResumenIvK
End Sub

Private Sub mnurevindice_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#End If
    
    frmRevisaIndice.Inicio
End Sub

Private Sub mnuSaldoIV_Click()
    frmIVSI.Show
End Sub

Private Sub mnuTotalVentas_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As Form
    
    Set f = BuscaForm("frmTotalVentas", "")
    If f Is Nothing Then Set f = New frmTotalVentas
    f.Inicio "VxTrans"
#End If
End Sub

Private Sub mnuTotalVentasProm_Click()
'jeaa 21/07/2009
    Dim f As Form
    
    Set f = BuscaForm("frmTotalVentas", "")
    If f Is Nothing Then Set f = New frmTotalVentas
    f.InicioPromVentas "VxTransProm"

End Sub

Private Sub mnuTransErradas_Click()
'jeaa 21/02/2006
    frmImprimePorLote.InicioPerdidaIdAsignado
End Sub

Private Sub mnuTransExportar_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As Form
    
    Set f = BuscaForm("frmExportar", "")
    If f Is Nothing Then Set f = New frmExportar
    f.Inicio
#End If
End Sub

Private Sub mnuTransImportar_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As Form
    
    Set f = BuscaForm("frmImportar", "")
    If f Is Nothing Then Set f = New frmImportar
    f.Inicio
#End If
End Sub

Private Sub mnuTransporte_Click()
frmFacturarxLote.InicioxGarante
End Sub

Private Sub mnuVentas3m_Click()
    Dim f As frm3M
    Set f = BuscaForm("frm3M", "")
    If f Is Nothing Then Set f = New frm3M
    f.InicioVentas ""
End Sub

Private Sub mnuVentasLocutorios_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmImportacion
    
    Set f = BuscaForm("frmImportacion", "VENTASLOCUTORIOS")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "VENTASLOCUTORIOS"
#End If
End Sub

Private Sub mnuVerificaCIRUC_Click()
'En versión DAO no está soportada la función
'jeaa 24/09/04 asignacion de grupo a los items
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    Set f = BuscaForm("frmITEMIVA_Cuenta", "PCCLIRUC")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "PCCLIRUC"
#End If
    
End Sub

Private Sub mnuVerificaCIRUCFactEelc_Click()
'En versión DAO no está soportada la función
'jeaa 24/09/04 asignacion de grupo a los items
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    Set f = BuscaForm("frmITEMIVA_Cuenta", "PCCLIRUCFCEL")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "PCCLIRUCFCEL"
#End If

End Sub

Private Sub mnuVerificaemail_Click()
'En versión DAO no está soportada la función
'jeaa 24/09/04 asignacion de grupo a los items
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    Set f = BuscaForm("frmITEMIVA_Cuenta", "PCEMAIL")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "PCEMAIL"
#End If
End Sub

Private Sub mnuVerificaITEMFactEelc_Click()
'En versión DAO no está soportada la función
'jeaa 24/09/04 asignacion de grupo a los items
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    Set f = BuscaForm("frmITEMIVA_Cuenta", "ITEMFCEL")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "ITEMFCEL"
#End If

End Sub

Private Sub mnuVerificaNotaEntrega_Click()
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As Form
    
    Set f = BuscaForm("NotaEntrega", "")
    If f Is Nothing Then Set f = New frmRegeneraTransformacion
    f.InicioNotaEntrega "NotaEntrega"
#End If

End Sub

Private Sub mnuVerificaPagos_Click()
    frmImprimePorLote.InicioVerificaPagosErrados
End Sub

Private Sub mnuVerificaTransformacion_Click()
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As Form
    
    Set f = BuscaForm("Transformacion", "")
    If f Is Nothing Then Set f = New frmRegeneraTransformacion
    f.Inicio
#End If
End Sub

'Private Sub mnuWindowArrangeIcons_Click()
''    Me.Arrange vbArrangeIcons
'End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub


Private Sub mnuFileExit_Click()
    'Descarga el formulario
    Unload Me
End Sub

'***Desabilitado hasta encontrar una manera de controlar cambios
'***Angel. 23/Marzo/2004
Public Function RecuperaRegistroIVFisico() As Boolean
'    On Error GoTo Errtrap
'
'    RecuperaRegistroIVFisico = False
'
'    With reg1
'        .hKey = HKEY_CLASSES_ROOT
'        .SubKeyPath = APPNAME_HIDE
'        .ValueName = "Validate"
'        .GetValue
'        If .value = "1978" Then
'            RecuperaRegistroIVFisico = True
'        End If
'    End With
'    Exit Function
'
'Errtrap:
'    RecuperaRegistroIVFisico = False
End Function

Private Sub mnucierreNoDev_Click()
'AUC 22/05/08 para transferir alquileres que no se develven todavia a otra empresa
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#End If
    frmCierreTransXDevolver.Inicio
End Sub

Private Sub mnuRegeneraRecetas_Click()
    frmRegeneraRecetas.Inicio
End Sub

Private Sub mnuGenHistorial_Click()
    frmHistorialPcProvcli.Inicio
End Sub

    
Private Sub mnuCostosnew_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemCostosNew
    
    Set f = BuscaForm("frmItemCostos", "Costos")
    If f Is Nothing Then Set f = New frmItemCostosNew
    f.Inicio "Costos"
    
#End If

End Sub

Private Sub mnuCostos1n_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemCostosNew
    
    Set f = BuscaForm("frmItemCostos", "Produccion")
    If f Is Nothing Then Set f = New frmItemCostosNew
    f.Inicio "Produccion"
    
#End If
End Sub

Private Sub mnuIVGrupo1_Click()
    Dim f As frmImportacion
    Set f = BuscaForm("frmImportacion", "IVGRUPO1")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "IVGRUPO1"
End Sub

Private Sub mnuIVGrupo2_Click()
    Dim f As frmImportacion
    Set f = BuscaForm("frmImportacion", "IVGRUPO2")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "IVGRUPO2"

End Sub

Private Sub mnuIVGrupo3_Click()
    Dim f As frmImportacion
    Set f = BuscaForm("frmImportacion", "IVGRUPO3")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "IVGRUPO3"
End Sub


Private Sub mnuIVGrupo4_Click()
    Dim f As frmImportacion
    Set f = BuscaForm("frmImportacion", "IVGRUPO4")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "IVGRUPO4"
End Sub

Private Sub mnuIVGrupo5_Click()
    Dim f As frmImportacion
    Set f = BuscaForm("frmImportacion", "IVGRUPO5")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "IVGRUPO5"
End Sub

Private Sub mnuAFGrupo1_Click()
    Dim f As frmImportacion
    Set f = BuscaForm("frmImportacion", "AFGRUPO1")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "AFGRUPO1"
End Sub

Private Sub mnuAFGrupo2_Click()
    Dim f As frmImportacion
    Set f = BuscaForm("frmImportacion", "AFGRUPO2")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "AFGRUPO2"

End Sub

Private Sub mnuAFGrupo3_Click()
    Dim f As frmImportacion
    Set f = BuscaForm("frmImportacion", "AFGRUPO3")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "AFGRUPO3"
End Sub


Private Sub mnuAFGrupo4_Click()
    Dim f As frmImportacion
    Set f = BuscaForm("frmImportacion", "AFGRUPO4")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "AFGRUPO4"
End Sub

Private Sub mnuAFGrupo5_Click()
    Dim f As frmImportacion
    Set f = BuscaForm("frmImportacion", "AFGRUPO5")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "AFGRUPO5"
End Sub


Private Sub mnuVendedores_Click()

'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmItemIVA_Cuenta
    Set f = BuscaForm("frmITEMIVA_Cuenta", "VENDE")
    If f Is Nothing Then Set f = New frmItemIVA_Cuenta
    f.Inicio "VENDE"
#End If
End Sub

Private Sub mnuEmpleado_Click()

'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmImportacion
    Set f = BuscaForm("frmImportacion", "PCEMP")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "PCEMP"
#End If
End Sub

Private Sub mnuactualComisionItem_Click()
'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmComisionesVendedor
    Set f = BuscaForm("frmComisionesVendedor", "COMIXITEM")
    If f Is Nothing Then Set f = New frmComisionesVendedor
    f.Inicio "COMIXITEM"
#End If
End Sub


Private Sub mnuCambiaAcopiadorPorLote_Click()
    'jeaa 01/10/2009
    frmImprimePorLote.InicioAcopiador
End Sub

Private Sub mnuImportarNumSerie_Click()

'En versión DAO no está soportada la función
#If DAOLIB Then
    MsgBox MSGERR_NODAO, vbInformation
    Exit Sub
#Else
    Dim f As frmImportacion
    Set f = BuscaForm("frmImportacion", "INVENTARIOSERIES")
    If f Is Nothing Then Set f = New frmImportacion
    f.Inicio "INVENTARIOSERIES"
#End If
End Sub
