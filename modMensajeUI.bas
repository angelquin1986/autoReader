Attribute VB_Name = "modMensaje"
Option Explicit

'Los mensajes comunes
Public Const MSG_PREPARA = "Está preparando..."
Public Const MSG_GRABANDO = "Está grabando..."
Public Const MSG_GENERANDOASIENTO = "Está generando asiento..."
Public Const MSG_BUSCANDO = "Está buscando..."
Public Const MSG_ACTUALIZA = "Está actualizando..."
Public Const MSG_NODISPONE = "La pantalla actual no dispone de la función."
Public Const MSG_CANCELMOD = "El registro está modificado." & vbCr & _
                                    "Desea grabar la modificación?"

'Public Const MSG_PREGUNTAREPORTE1 = "Desea imprimir el comprobante?"
Public Const MSG_PREGUNTAREPORTE2 = "Si desea imprimir el comprobante, confirme que la impresora está lista y apláste el botón 'Sí'."
Public Const MSG_ERR_FECHA = "La fecha está mal."
Public Const MSG_ERR_SINDETALLE = "No se puede grabar sin detalle."
Public Const MSG_ERR_RESPONSABLE = "Seleccione el responsable."
Public Const MSG_ERR_NOGRABA = "No se pudo grabar el comprobante."
Public Const MSG_ERR_NOVISUALIZA = _
                "La transacción seleccionada no se dispone de visualizar en ésta función." & vbCr & _
                "Se puede hacerlo en el módulo correspondiente."
Public Const MSG_ERR_REPITE_NUMTRANS = _
                "No se puede grabar el comprobante debido a que " & _
                "ya existe el mismo numero. Por favor cambie el " & _
                "numero de comprobante e intente de nuevo."
Public Const MSG_ERR_MAYORALSALDO = "El valor no puede ser mayor al saldo."
Public Const MSG_ERR_BASETEMPNO = "No se puede abrir la base temporal."
Public Const MSG_ERR_NOPERMITEELIMINAR = "No se puede eliminar/anular ésta transacción debido a que existe otros comprobantes sobre ésta."
Public Const MSG_ERR_NOPERMITEMOD = "No se puede modificar ésta transaccion debido a que existe otros comprobantes sobre ésta."
Public Const MSG_ERR_NOENCUENTRA = "No se encuentra el código. "

