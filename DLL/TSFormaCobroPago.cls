VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TSFormaCobroPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Empresa"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type T_PROPS
    IDForma As Long
    CodForma As String
    NombreForma As String
    Plazo As Integer
    CambiaFechaVenci As Boolean
    PermiteAbono As Boolean
    BandCobro As Boolean
    BandValida As Boolean
    BandEliminable As Boolean
    IdBanco As Long
    CodBanco As String
    IdTipoDoc As Long
    CodTipoDoc As String
    NotaCredito As String 'Agregado AUC 12/10/2005
    
    ConsiderarComoEfectivo As Boolean  'Agregado para poder distinguir si la  Forma de CobroPago es efectivo
                                       'en caso de tarjetas de credito, no van directo al Banco pero son Efectivo
    IngresoAutomatico As Boolean     'Ingreso automatico con cambio de cliente
    IdProvcli As Long
    CodProvcli As String
    IdFormaTC As Long
    CodFormaTC As String
    DeudaMismoCliente As Boolean
    PorDesc As Double   'jeaa 10/10/2008 mp3 para desto x forma cobro
    ConsiderarComoNegativo As Boolean 'Agregado AUC 22/01/2007
    DatosAdicionales As Boolean  ' jeaa 17/07/2009
    ControlPrecios As Boolean  ' jeaa 17/07/2009
    ListaPrecios As String
                                        
    DeudaClienteIVBanco As Boolean
    DatosObligatorios As String
    NumLetra As Integer
    
    NuevoPlazo As Integer
    IdCuentaActivo As Long
    IdCuentaVenta As Long
    IdCuentaDiferida As Long
    CodCuentaActivo As String
    CodCuentaVenta As String
    CodCuentaDiferida As String
    NoAplicaDiferidos As Boolean
    BandIncNumLetra As Boolean
    NumLetraNew As Long
    
    IdFormaSRI As Long
    CodFormaSRI As String
    
    
    fechagrabado As Date
    objEmpresa As Empresa       'Referencia al objeto padre
    
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Public Sub Grabar()
    Dim sql As String, rs As Recordset
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
    sql = "SELECT CodForma FROM TSFormaCobroPago " & _
          "WHERE (CodForma='" & mProps.CodForma & "') AND " & _
                "(IdForma<>" & mProps.IDForma & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "TSFormaCobroPago", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM TSFormaCobroPago WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM TSFormaCobroPago WHERE IdForma=" & mProps.IDForma
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "TSFormaCobroPago_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en Pos
    With rs
        !CodForma = mProps.CodForma
        !NombreForma = mProps.NombreForma
        !Plazo = mProps.Plazo
        !CambiaFechaVenci = mProps.CambiaFechaVenci
        !PermiteAbono = mProps.PermiteAbono
        !BandCobro = mProps.BandCobro
        !BandValida = mProps.BandValida
        !BandEliminable = mProps.BandEliminable
        !IdBanco = mProps.IdBanco
        !IdTipoDoc = mProps.IdTipoDoc
        !ConsiderarComoEfectivo = mProps.ConsiderarComoEfectivo   'agregado Oliver 3/oct/2002
        !IdProvcli = mProps.IdProvcli
        !IdFormaTC = mProps.IdFormaTC
        !IngresoAutomatico = mProps.IngresoAutomatico 'jeaa 16/04/2008
        !DeudaMismoCliente = mProps.DeudaMismoCliente 'jeaa 16/04/2008
        !fechagrabado = Now
        !NotaCredito = mProps.NotaCredito 'Agreagdo AUC 13/10/2005
        !PorDesc = mProps.PorDesc   'jeaa 10/10/2008 mp3
        !ConsiderarComoNegativo = mProps.ConsiderarComoNegativo 'AUC 22/01/07
        !DatosAdicionales = mProps.DatosAdicionales 'jeaa 17/07/2009
        !ControlPrecios = mProps.ControlPrecios 'jeaa 17/07/2009
        !ListaPrecios = mProps.ListaPrecios 'jeaa 17/07/2009
        !DeudaClienteIVBanco = mProps.DeudaClienteIVBanco
        !DatosObligatorios = mProps.DatosObligatorios
        '!NumLetra = mProps.NumLetra
        !NumLetra = mProps.NumLetraNew
        !NuevoPlazo = mProps.NuevoPlazo
        !IdCuentaActivo = mProps.IdCuentaActivo
        !IdCuentaVenta = mProps.IdCuentaVenta
        !IdCuentaDiferida = mProps.IdCuentaDiferida
        !NoAplicaDiferidos = mProps.NoAplicaDiferidos
        !BandIncNumLetra = mProps.BandIncNumLetra
        !IdFormaSRI = mProps.IdFormaSRI

        !fechagrabado = mProps.fechagrabado
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IDForma = !IDForma
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_FCP", "Forma Cobro/Pago Nueva: " & _
                                                         mProps.CodForma, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_FCP", "Forma Cobro/Pago Modificada: " & _
                                                         mProps.CodForma, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_FORMACOBROPAGO
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM vwTSForma WHERE CodForma='" & key & "'"
    Else
        sql = "SELECT * FROM vwTSForma WHERE IdForma=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IDForma = !IDForma
        mProps.CodForma = !CodForma
        If Not IsNull(!NombreForma) Then mProps.NombreForma = !NombreForma
        If Not IsNull(!Plazo) Then mProps.Plazo = !Plazo
        If Not IsNull(!CambiaFechaVenci) Then mProps.CambiaFechaVenci = !CambiaFechaVenci
        If Not IsNull(!PermiteAbono) Then mProps.PermiteAbono = !PermiteAbono
        If Not IsNull(!BandCobro) Then mProps.BandCobro = !BandCobro
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!BandEliminable) Then mProps.BandEliminable = !BandEliminable
        If Not IsNull(!IdBanco) Then mProps.IdBanco = !IdBanco
        If Not IsNull(!CodBanco) Then mProps.CodBanco = !CodBanco
        If Not IsNull(!IdTipoDoc) Then mProps.IdTipoDoc = !IdTipoDoc
        If Not IsNull(!CodTipoDoc) Then mProps.CodTipoDoc = !CodTipoDoc
        If Not IsNull(!ConsiderarComoEfectivo) Then mProps.ConsiderarComoEfectivo = !ConsiderarComoEfectivo  'Agregado por Oliver
        If Not IsNull(!NotaCredito) Then mProps.NotaCredito = !NotaCredito
        If Not IsNull(!IdProvcli) Then mProps.IdProvcli = !IdProvcli
        If Not IsNull(!CodProvcli) Then mProps.CodProvcli = !CodProvcli
        If Not IsNull(!IdFormaTC) Then mProps.IdFormaTC = !IdFormaTC
        If Not IsNull(!CodFormaTC) Then mProps.CodFormaTC = !CodFormaTC
        If Not IsNull(!IngresoAutomatico) Then mProps.IngresoAutomatico = !IngresoAutomatico 'Agregado por jeaa 16/04/2008
        If Not IsNull(!DeudaMismoCliente) Then mProps.DeudaMismoCliente = !DeudaMismoCliente 'Agregado por jeaa 16/04/2008
        If Not IsNull(!PorDesc) Then mProps.PorDesc = !PorDesc '10/10/2008 mp3
        If Not IsNull(!ConsiderarComoNegativo) Then mProps.ConsiderarComoNegativo = !ConsiderarComoNegativo
        If Not IsNull(!DatosAdicionales) Then mProps.DatosAdicionales = !DatosAdicionales 'Agregado por jeaa 17/07/2009
        If Not IsNull(!ControlPrecios) Then mProps.ControlPrecios = !ControlPrecios 'Agregado por jeaa 17/07/2009
        If Not IsNull(!ListaPrecios) Then mProps.ListaPrecios = !ListaPrecios 'Agregado por jeaa 17/07/2009
        If Not IsNull(!DeudaClienteIVBanco) Then mProps.DeudaClienteIVBanco = !DeudaClienteIVBanco 'Agregado por jeaa 16/04/2008
        If Not IsNull(!DatosObligatorios) Then mProps.DatosObligatorios = !DatosObligatorios 'Agregado por jeaa 14/12/2011
        'If Not IsNull(!NumLetra) Then mProps.NumLetra = !NumLetra
        If Not IsNull(!NumLetra) Then mProps.NumLetraNew = !NumLetra
        If Not IsNull(!NuevoPlazo) Then mProps.NuevoPlazo = !NuevoPlazo
        If Not IsNull(!IdCuentaActivo) Then mProps.IdCuentaActivo = !IdCuentaActivo
        If Not IsNull(!IdCuentaVenta) Then mProps.IdCuentaVenta = !IdCuentaVenta
        If Not IsNull(!IdCuentaDiferida) Then mProps.IdCuentaDiferida = !IdCuentaDiferida
        If Not IsNull(!CodCuentaActivo) Then mProps.CodCuentaActivo = !CodCuentaActivo
        If Not IsNull(!CodCuentaVenta) Then mProps.CodCuentaVenta = !CodCuentaVenta
        If Not IsNull(!CodCuentaDiferida) Then mProps.CodCuentaDiferida = !CodCuentaDiferida
        If Not IsNull(!NoAplicaDiferidos) Then mProps.NoAplicaDiferidos = !NoAplicaDiferidos
        mProps.BandIncNumLetra = !BandIncNumLetra
        If Not IsNull(!IdFormaSRI) Then mProps.IdFormaSRI = !IdFormaSRI
        If Not IsNull(!CodFormaSRI) Then mProps.CodFormaSRI = !CodFormaSRI

        
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset

    'Verifica si existe Kardex
    sql = "SELECT IdForma FROM PCKardex WHERE IdForma=" & mProps.IDForma
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "TSFormaCobroPago.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing

    'Si no es eliminable
    If Not mProps.BandEliminable Then
        Err.Raise ERR_NOELIMINABLE, "TSFormaCobroPago_Eliminar", MSGERR_NOELIMINABLE
    End If

    'Elimina la forma de cobro/pago
    sql = "DELETE FROM TSFormaCobroPago WHERE CodForma='" & mProps.CodForma & "'"
    mProps.objEmpresa.Execute sql, True

    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_FCP", "Forma de Cobro/Pago Eliminada: " & _
                                                         mProps.CodForma, "DT"
    
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_FORMACOBROPAGO
End Sub

Public Function Clone() As TSFormaCobroPago
    Dim nuevo As TSFormaCobroPago
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New TSFormaCobroPago
    nuevo.Props = mProps
    nuevo.BandEliminable = True
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IDForma = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property



Public Property Get IDForma() As Long
    IDForma = mProps.IDForma
End Property

Public Property Let CodForma(ByVal value As String)
    If Len(value) > 5 Then
        Err.Raise ERR_INVALIDO, "TSFormaCobroPago_CodForma", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodForma = value
    mbooModificado = True
End Property

Public Property Get CodForma() As String
    CodForma = Trim$(mProps.CodForma)
End Property

'AUC 13/10/2005
Public Property Let NotaCredito(ByVal value As String)
 If Len(value) > 120 Then
     Err.Raise ERR_INVALIDO, "TSFormaCobroPabo_NotaCredito", MSGERR_INVALIDO
 End If
 mProps.NotaCredito = value
 mbooModificado = True
End Property

'AUC 13/10/2005
Public Property Get NotaCredito() As String
    NotaCredito = Trim$(mProps.NotaCredito)
End Property

Public Property Let NombreForma(ByVal value As String)
    If Len(value) > 100 Then
        Err.Raise ERR_INVALIDO, "TSFormaCobroPago_NombreForma", MSGERR_INVALIDO
    End If
    
    mProps.NombreForma = value
    mbooModificado = True
End Property

Public Property Get NombreForma() As String
    NombreForma = Trim$(mProps.NombreForma)
End Property

Public Property Let Plazo(ByVal value As Integer)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "TSFormaCobroPago_Plazo", MSGERR_INVALIDO
    End If
    
    mProps.Plazo = value
    mbooModificado = True
End Property

Public Property Get Plazo() As Integer
    Plazo = mProps.Plazo
End Property

Public Property Let CambiaFechaVenci(ByVal value As Boolean)
    mProps.CambiaFechaVenci = value
    mbooModificado = True
End Property

Public Property Get CambiaFechaVenci() As Boolean
    CambiaFechaVenci = mProps.CambiaFechaVenci
End Property

Public Property Let PermiteAbono(ByVal value As Boolean)
    mProps.PermiteAbono = value
    mbooModificado = True
End Property

Public Property Get PermiteAbono() As Boolean
    PermiteAbono = mProps.PermiteAbono
End Property

Public Property Let BandCobro(ByVal value As Boolean)
    mProps.BandCobro = value
    mbooModificado = True
End Property

Public Property Get BandCobro() As Boolean
    BandCobro = mProps.BandCobro
End Property


Public Property Let ConsiderarComoEfectivo(ByVal value As Boolean)
    mProps.ConsiderarComoEfectivo = value
    mbooModificado = True
End Property

Public Property Get ConsiderarComoEfectivo() As Boolean
    ConsiderarComoEfectivo = mProps.ConsiderarComoEfectivo
End Property

Public Property Let BandValida(ByVal value As Boolean)
    mProps.BandValida = value
    mbooModificado = True
End Property

Public Property Get BandValida() As Boolean
    BandValida = mProps.BandValida
End Property

Public Property Let BandEliminable(ByVal value As Boolean)
    mProps.BandEliminable = value
    mbooModificado = True
End Property

Public Property Get BandEliminable() As Boolean
    BandEliminable = mProps.BandEliminable
End Property

Public Property Let IdBanco(ByVal value As Long)
    mProps.IdBanco = value
    mbooModificado = True
End Property

Public Property Get IdBanco() As Long
    IdBanco = mProps.IdBanco
End Property


Public Property Let CodBanco(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodBanco Then
        If Len(value) > 0 Then
            'Actualiza IdBanco también
            sql = "SELECT IdBanco FROM TSBanco WHERE CodBanco='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "TSKardex_CodBanco", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdBanco = rs!IdBanco
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdBanco = 0
        End If
        
        mProps.CodBanco = value
        mbooModificado = True
    End If
End Property

Public Property Get CodBanco() As String
    CodBanco = Trim$(mProps.CodBanco)
End Property


Public Property Let IdTipoDoc(ByVal value As Long)
    mProps.IdTipoDoc = value
    mbooModificado = True
End Property

Public Property Get IdTipoDoc() As Long
    IdTipoDoc = mProps.IdTipoDoc
End Property


Public Property Let CodTipoDoc(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodTipoDoc Then
        If Len(value) > 0 Then
            'Actualiza IdTipoDoc también
            sql = "SELECT IdTipoDoc FROM TSTipoDocBanco WHERE CodTipoDoc='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "TSKardex_CodTipoDoc", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdTipoDoc = rs!IdTipoDoc
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdTipoDoc = 0
        End If
        
        mProps.CodTipoDoc = value
        mbooModificado = True
    End If
End Property

Public Property Get CodTipoDoc() As String
    CodTipoDoc = Trim$(mProps.CodTipoDoc)
End Property

Public Property Get fechagrabado() As Date
    fechagrabado = mProps.fechagrabado
End Property

Public Property Get EsNuevo() As Boolean
    EsNuevo = mbooEsNuevo
End Property

Public Property Get Modificado() As Boolean
    Modificado = mbooModificado
End Property


Private Sub Class_Initialize()
    With mProps
        .BandCobro = True
        .BandEliminable = True
        .BandValida = True
        .CambiaFechaVenci = True
        .PermiteAbono = True
        .IngresoAutomatico = False
        .Plazo = 0
        .PorDesc = 0
        .NuevoPlazo = 0
    End With
    
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Property Let IdProvcli(ByVal value As Long)
    mProps.IdProvcli = value
    mbooModificado = True
End Property

Public Property Get IdProvcli() As Long
    IdProvcli = mProps.IdProvcli
End Property


Public Property Let CodProvcli(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodProvcli Then
        If Len(value) > 0 Then
            'Actualiza IdProvCli también
            sql = "SELECT IdProvCli FROM PcProvCli WHERE bandcliente=1 and CodProvCli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "TSKardex_CodProvCli", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdProvcli = rs!IdProvcli
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdProvcli = 0
        End If
        
        mProps.CodProvcli = value
        mbooModificado = True
    End If
End Property

Public Property Get CodProvcli() As String
    CodProvcli = Trim$(mProps.CodProvcli)
End Property


Public Property Let IdFormaTC(ByVal value As Long)
    mProps.IdFormaTC = value
    mbooModificado = True
End Property


Public Property Get IdFormaTC() As Long
    IdFormaTC = mProps.IdFormaTC
End Property


Public Property Let CodFormaTC(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodFormaTC Then
        If Len(value) > 0 Then
            'Actualiza IdFormaTC también
            sql = "SELECT IdForma FROM TSFormaCobroPago WHERE CodForma='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "TSKardex_CodFormaTC", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdFormaTC = rs!IDForma
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdFormaTC = 0
        End If
        
        mProps.CodFormaTC = value
        mbooModificado = True
    End If
End Property


Public Property Get CodFormaTC() As String
    CodFormaTC = Trim$(mProps.CodFormaTC)
End Property


Public Property Let IngresoAutomatico(ByVal value As Boolean)
    mProps.IngresoAutomatico = value
    mbooModificado = True
End Property

Public Property Get IngresoAutomatico() As Boolean
    IngresoAutomatico = mProps.IngresoAutomatico
End Property

Public Property Let DeudaMismoCliente(ByVal value As Boolean)
    mProps.DeudaMismoCliente = value
    mbooModificado = True
End Property

Public Property Get DeudaMismoCliente() As Boolean
    DeudaMismoCliente = mProps.DeudaMismoCliente
End Property

Public Property Let PorDesc(ByVal value As Double)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "TSFormaCobroPago_PorDesc", MSGERR_INVALIDO
    End If
    
    mProps.PorDesc = value
    mbooModificado = True
End Property

Public Property Get PorDesc() As Double
    PorDesc = mProps.PorDesc
End Property

'AUC 21/01/07
Public Property Let ConsiderarComoNegativo(ByVal value As Boolean)
    mProps.ConsiderarComoNegativo = value
    mbooModificado = True
End Property

Public Property Get ConsiderarComoNegativo() As Boolean
    ConsiderarComoNegativo = mProps.ConsiderarComoNegativo
End Property


Public Property Let DatosAdicionales(ByVal value As Boolean)
    mProps.DatosAdicionales = value
    mbooModificado = True
End Property

Public Property Get DatosAdicionales() As Boolean
    DatosAdicionales = mProps.DatosAdicionales
End Property

Public Property Let ControlPrecios(ByVal value As Boolean)
    mProps.ControlPrecios = value
    mbooModificado = True
End Property

Public Property Get ControlPrecios() As Boolean
    ControlPrecios = mProps.ControlPrecios
End Property

Public Property Let ListaPrecios(ByVal value As String)
    mProps.ListaPrecios = value
    mbooModificado = True
End Property

Public Property Get ListaPrecios() As String
    ListaPrecios = Trim$(mProps.ListaPrecios)
End Property


Public Property Let DeudaClienteIVBanco(ByVal value As Boolean)
    mProps.DeudaClienteIVBanco = value
    mbooModificado = True
End Property

Public Property Get DeudaClienteIVBanco() As Boolean
    DeudaClienteIVBanco = mProps.DeudaClienteIVBanco
End Property

Public Property Let DatosObligatorios(ByVal value As String)
    If Len(value) > 8 Then
        Err.Raise ERR_INVALIDO, "TSFormaCobroPago_DatosObligatorios", MSGERR_INVALIDO
    End If
    
    mProps.DatosObligatorios = value
    mbooModificado = True
End Property

Public Property Get DatosObligatorios() As String
    DatosObligatorios = Trim$(mProps.DatosObligatorios)
End Property

Public Property Let NumLetra(ByVal value As Integer)
'    If value < 0 Then
'        Err.Raise ERR_INVALIDO, "TSFormaCobroPago_NumLetra", MSGERR_INVALIDO
'    End If
'    mProps.NumLetra = value
'    mbooModificado = True
End Property

Public Property Get NumLetra() As Integer
'    NumLetra = mProps.NumLetra
End Property

Public Property Let NuevoPlazo(ByVal value As Integer)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "TSFormaCobroPago_NuevoPlazo", MSGERR_INVALIDO
    End If
    
    mProps.NuevoPlazo = value
    mbooModificado = True
End Property

Public Property Get NuevoPlazo() As Integer
    NuevoPlazo = mProps.NuevoPlazo
End Property

Public Property Let IdCuentaActivo(ByVal value As Long)
    mProps.IdCuentaActivo = value
   mbooModificado = True
End Property
Public Property Get IdCuentaActivo() As Long
    IdCuentaActivo = mProps.IdCuentaActivo
End Property
Public Property Let CodCuentaActivo(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodCuentaActivo Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVInventario_CodCuentaActivo", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaActivo = rs!IdCuenta
           End If
            rs.Close
        Else
            IdCuentaActivo = 0
        End If
        mProps.CodCuentaActivo = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property
Public Property Get CodCuentaActivo() As String
    CodCuentaActivo = Trim$(mProps.CodCuentaActivo)
End Property
Public Property Let IdCuentaVenta(ByVal value As Long)
    mProps.IdCuentaVenta = value
    mbooModificado = True
End Property
Public Property Get IdCuentaVenta() As Long
    IdCuentaVenta = mProps.IdCuentaVenta
End Property
Public Property Let CodCuentaVenta(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodCuentaVenta Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaVenta también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVInventario_CodCuentaVenta", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaVenta = rs!IdCuenta
            End If
            rs.Close
        Else
            IdCuentaVenta = 0
        End If
        mProps.CodCuentaVenta = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property
Public Property Get CodCuentaVenta() As String
    CodCuentaVenta = Trim$(mProps.CodCuentaVenta)
End Property
Public Property Let IdCuentaDiferida(ByVal value As Long)
    mProps.IdCuentaDiferida = value
    mbooModificado = True
End Property
Public Property Get IdCuentaDiferida() As Long
    IdCuentaDiferida = mProps.IdCuentaDiferida
End Property
Public Property Let CodCuentaDiferida(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodCuentaDiferida Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaDiferida también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVInDiferidario_CodCuentaDiferida", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaDiferida = rs!IdCuenta
            End If
            rs.Close
        Else
            IdCuentaDiferida = 0
        End If
        mProps.CodCuentaDiferida = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property

Public Property Get CodCuentaDiferida() As String
    CodCuentaDiferida = Trim$(mProps.CodCuentaDiferida)
End Property

Public Property Let NoAplicaDiferidos(ByVal value As Boolean)
    mProps.NoAplicaDiferidos = value
    mbooModificado = True
End Property

Public Property Get NoAplicaDiferidos() As Boolean
    NoAplicaDiferidos = mProps.NoAplicaDiferidos
End Property


Public Property Let NumLetraNew(ByVal value As Long)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "TSFormaCobroPago_NumLetraNew", MSGERR_INVALIDO
    End If
    mProps.NumLetraNew = value
    mbooModificado = True
End Property

Public Property Get NumLetraNew() As Long
    NumLetraNew = mProps.NumLetraNew
End Property
    
Public Property Let BandIncNumLetra(ByVal value As Boolean)
    mProps.BandIncNumLetra = value
    mbooModificado = True
End Property

Public Property Get BandIncNumLetra() As Boolean
    BandIncNumLetra = mProps.BandIncNumLetra
End Property

Public Property Let IdFormaSRI(ByVal value As Long)
    mProps.IdFormaSRI = value
    mbooModificado = True
End Property

Public Property Get IdFormaSRI() As Long
    IdFormaSRI = mProps.IdFormaSRI
End Property


Public Property Let CodFormaSRI(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodFormaSRI Then
        If Len(value) > 0 Then
            'Actualiza IdFormaSRI también
            sql = "SELECT Id FROM Anexo_FormaPago WHERE CodFormaPago='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "TSKardex_CodFormaSRI", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdFormaSRI = rs!id
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdFormaSRI = 0
        End If
        
        mProps.CodFormaSRI = value
        mbooModificado = True
    End If
End Property

Public Property Get CodFormaSRI() As String
    CodFormaSRI = Trim$(mProps.CodFormaSRI)
End Property

