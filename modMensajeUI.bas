Attribute VB_Name = "modMensaje"
Option Explicit

'Los mensajes comunes
Public Const MSG_PREPARA = "Est� preparando..."
Public Const MSG_GRABANDO = "Est� grabando..."
Public Const MSG_GENERANDOASIENTO = "Est� generando asiento..."
Public Const MSG_BUSCANDO = "Est� buscando..."
Public Const MSG_ACTUALIZA = "Est� actualizando..."
Public Const MSG_NODISPONE = "La pantalla actual no dispone de la funci�n."
Public Const MSG_CANCELMOD = "El registro est� modificado." & vbCr & _
                                    "Desea grabar la modificaci�n?"

'Public Const MSG_PREGUNTAREPORTE1 = "Desea imprimir el comprobante?"
Public Const MSG_PREGUNTAREPORTE2 = "Si desea imprimir el comprobante, confirme que la impresora est� lista y apl�ste el bot�n 'S�'."
Public Const MSG_ERR_FECHA = "La fecha est� mal."
Public Const MSG_ERR_SINDETALLE = "No se puede grabar sin detalle."
Public Const MSG_ERR_RESPONSABLE = "Seleccione el responsable."
Public Const MSG_ERR_NOGRABA = "No se pudo grabar el comprobante."
Public Const MSG_ERR_NOVISUALIZA = _
                "La transacci�n seleccionada no se dispone de visualizar en �sta funci�n." & vbCr & _
                "Se puede hacerlo en el m�dulo correspondiente."
Public Const MSG_ERR_REPITE_NUMTRANS = _
                "No se puede grabar el comprobante debido a que " & _
                "ya existe el mismo numero. Por favor cambie el " & _
                "numero de comprobante e intente de nuevo."
Public Const MSG_ERR_MAYORALSALDO = "El valor no puede ser mayor al saldo."
Public Const MSG_ERR_BASETEMPNO = "No se puede abrir la base temporal."
Public Const MSG_ERR_NOPERMITEELIMINAR = "No se puede eliminar/anular �sta transacci�n debido a que existe otros comprobantes sobre �sta."
Public Const MSG_ERR_NOPERMITEMOD = "No se puede modificar �sta transaccion debido a que existe otros comprobantes sobre �sta."
Public Const MSG_ERR_NOENCUENTRA = "No se encuentra el c�digo. "

