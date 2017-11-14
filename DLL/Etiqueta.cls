VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Etiqueta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type T_POSCAMPO
    X As Single
    Y As Single
    Fuente As String
    TamLetra As Integer
    Imprimir As Boolean
    Longitud As Integer
    Formato As String
    Prefijo As String
End Type

Const MmTw = 56.6929
Const IND_EMP = 0 'Indice de datos de empresa
Const IND_COD = 1 'Indice de datos del código de item
Const IND_CA1 = 2 'Indice de datos del código ALTERNO de item
Const IND_DES = 3 'Indice de datos de la descripción del item
Const IND_COS = 4 'Indice de datos del costo
Const IND_PRE1 = 5 'Indice de datos del precio 1
Const IND_PRE2 = 6 'Indice de datos del precio 2
Const IND_PRE3 = 7 'Indice de datos del precio 3
Const IND_PRE4 = 8 'Indice de datos del precio 4
Const IND_PRE5 = 9 'Indice de datos del precio 5
Const IND_PRE6 = 10 'Indice de datos del precio 5
Const IND_PIVA1 = 11 '9 'Indice de datos del precio 1 + IVA
Const IND_PIVA2 = 12 '10 'Indice de datos del precio 2 + IVA
Const IND_PIVA3 = 13 '11 'Indice de datos del precio 3 + IVA
Const IND_PIVA4 = 14 '12 'Indice de datos del precio 4 + IVA
Const IND_PIVA5 = 15 '13 'Indice de datos del precio 5 + IVA
Const IND_PIVA6 = 16 '13 'Indice de datos del precio 6 + IVA AUC agregado precio6
Const IND_CBR = 17 '13 'Indice de datos del código de barras
Const IND_IVG1 = 18
Const IND_IVGD1 = 19
Const IND_IVG2 = 20
Const IND_IVGD2 = 21
Const IND_IVG3 = 22
Const IND_IVGD3 = 23
Const IND_IVG4 = 24
Const IND_IVGD4 = 25
Const IND_IVG5 = 26
Const IND_IVGD5 = 27
Const IND_CANT = 28
Const IND_CODDER = 29


Private mbooModificado As Boolean
Private MargenesCodBar As Variant

Private mPosCampo(30) As T_POSCAMPO   'Para definir posiciones dentro de cada etiqueta


Public Property Let Columnas(ByVal value As Integer)
'Valor predeterminado 5
    MargenesCodBar(0) = IIf(IsNumeric(value), value, 5)
    mbooModificado = True
End Property

Public Property Get Columnas() As Integer
    Columnas = MargenesCodBar(0)
End Property

Public Property Let Filas(ByVal value As Integer)
'Valor predeterminado 10
    MargenesCodBar(1) = IIf(IsNumeric(value), value, 10)
    mbooModificado = True
End Property

Public Property Get Filas() As Integer
    Filas = MargenesCodBar(1)
End Property

Public Property Let AltoPapel(ByVal value As Single)
'Valor predeterminado 290
    MargenesCodBar(2) = IIf(IsNumeric(value), value, 290)
    mbooModificado = True
End Property

Public Property Get AltoPapel() As Single
    AltoPapel = MargenesCodBar(2) * MmTw
End Property

Public Property Let AnchoPapel(ByVal value As Single)
'Valor predeterminado 210
    MargenesCodBar(3) = IIf(IsNumeric(value), value, 210)
    mbooModificado = True
End Property

Public Property Get AnchoPapel() As Single
    AnchoPapel = MargenesCodBar(3) * MmTw
End Property

Public Property Let MargenSuperior(ByVal value As Single)
'Valor predeterminado 20
    MargenesCodBar(4) = IIf(IsNumeric(value), value, 20)
    mbooModificado = True
End Property

Public Property Get MargenSuperior() As Single
    MargenSuperior = MargenesCodBar(4) * MmTw
End Property

Public Property Let MargenInferior(ByVal value As Single)
'Valor predeterminado 20
    MargenesCodBar(5) = IIf(IsNumeric(value), value, 20)
    mbooModificado = True
End Property

Public Property Get MargenInferior() As Single
    MargenInferior = MargenesCodBar(5) * MmTw
End Property

Public Property Let MargenIzquierdo(ByVal value As Single)
'Valor predeterminado 14
    MargenesCodBar(6) = IIf(IsNumeric(value), value, 14)
    mbooModificado = True
End Property

Public Property Get MargenIzquierdo() As Single
    MargenIzquierdo = MargenesCodBar(6) * MmTw
End Property

Public Property Let MargenDerecho(ByVal value As Single)
'Valor predeterminado 14
    MargenesCodBar(7) = IIf(IsNumeric(value), value, 14)
    mbooModificado = True
End Property

Public Property Get MargenDerecho() As Single
    MargenDerecho = MargenesCodBar(7) * MmTw
End Property

Public Property Let MargenImpresora(ByVal value As Single)
'Valor predeterminado 8
    MargenesCodBar(8) = IIf(IsNumeric(value), value, 8)
    mbooModificado = True
End Property

Public Property Get MargenImpresora() As Single
    MargenImpresora = MargenesCodBar(8) * MmTw
End Property

Public Property Let EspacioFilas(ByVal value As Single)
'Valor predeterminado 0.02
    MargenesCodBar(9) = IIf(IsNumeric(value), value, 0.02)
    mbooModificado = True
End Property

Public Property Get EspacioFilas() As Single
    EspacioFilas = MargenesCodBar(9) * MmTw
End Property

Public Property Let Bordes(ByVal value As Single)
'Valor predeterminado 0 = Sin Bordes; 1=Dibujar Bordes
    MargenesCodBar(10) = IIf(IsNumeric(value), value, 0)
    mbooModificado = True
End Property

Public Property Get Bordes() As Single
    Bordes = MargenesCodBar(10)
End Property

Public Property Get NumeroEtiquetas() As Integer
    NumeroEtiquetas = Filas * Columnas
End Property

'**********Propiedades para establecer las fuentes de los campos************
Public Property Get FuenteEmpresa() As String
    FuenteEmpresa = mPosCampo(IND_EMP).Fuente
End Property

Public Property Get FuenteCodigo() As String
    FuenteCodigo = mPosCampo(IND_COD).Fuente
End Property

Public Property Get FuenteCodAlterno() As String
    FuenteCodAlterno = mPosCampo(IND_CA1).Fuente
End Property

Public Property Get FuenteDescripcion() As String
    FuenteDescripcion = mPosCampo(IND_DES).Fuente
End Property

Public Property Get FuenteCosto() As String
    FuenteCosto = mPosCampo(IND_COS).Fuente
End Property

Public Property Get FuentePrecio1() As String
    FuentePrecio1 = mPosCampo(IND_PRE1).Fuente
End Property

Public Property Get FuentePrecio2() As String
    FuentePrecio2 = mPosCampo(IND_PRE2).Fuente
End Property

Public Property Get FuentePrecio3() As String
    FuentePrecio3 = mPosCampo(IND_PRE3).Fuente
End Property

Public Property Get FuentePrecio4() As String
    FuentePrecio4 = mPosCampo(IND_PRE4).Fuente
End Property
'AUC 29/11/07
Public Property Get FuentePrecio5() As String
    FuentePrecio5 = mPosCampo(IND_PRE5).Fuente
End Property
Public Property Get FuentePrecio6() As String
    FuentePrecio6 = mPosCampo(IND_PRE6).Fuente
End Property


Public Property Get FuentePrecioIVA1() As String
    FuentePrecioIVA1 = mPosCampo(IND_PIVA1).Fuente
End Property

Public Property Get FuentePrecioIVA2() As String
    FuentePrecioIVA2 = mPosCampo(IND_PIVA2).Fuente
End Property

Public Property Get FuentePrecioIVA3() As String
    FuentePrecioIVA3 = mPosCampo(IND_PIVA3).Fuente
End Property

Public Property Get FuentePrecioIVA4() As String
    FuentePrecioIVA4 = mPosCampo(IND_PIVA4).Fuente
End Property
Public Property Get FuentePrecioIVA5() As String
    FuentePrecioIVA5 = mPosCampo(IND_PIVA5).Fuente
End Property

Public Property Get FuentePrecioIVA6() As String
    FuentePrecioIVA6 = mPosCampo(IND_PIVA6).Fuente
End Property
Public Property Get FuenteCodigoBarra() As String
    FuenteCodigoBarra = mPosCampo(IND_CBR).Fuente
End Property

'**********Propiedades para establecer el tamaño de la letra de los campos************
Public Property Get TamLetraEmpresa() As Integer
    TamLetraEmpresa = mPosCampo(IND_EMP).TamLetra
End Property

Public Property Get TamLetraCodigo() As Integer
    TamLetraCodigo = mPosCampo(IND_COD).TamLetra
End Property

Public Property Get TamLetraCodAlterno() As Integer
    TamLetraCodAlterno = mPosCampo(IND_CA1).TamLetra
End Property

Public Property Get TamLetraDescripcion() As Integer
    TamLetraDescripcion = mPosCampo(IND_DES).TamLetra
End Property

Public Property Get TamLetraCosto() As Integer
    TamLetraCosto = mPosCampo(IND_COS).TamLetra
End Property

Public Property Get TamLetraPrecio1() As String
    TamLetraPrecio1 = mPosCampo(IND_PRE1).TamLetra
End Property

Public Property Get TamLetraPrecio2() As String
    TamLetraPrecio2 = mPosCampo(IND_PRE2).TamLetra
End Property

Public Property Get TamLetraPrecio3() As String
    TamLetraPrecio3 = mPosCampo(IND_PRE3).TamLetra
End Property

Public Property Get TamLetraPrecio4() As String
    TamLetraPrecio4 = mPosCampo(IND_PRE4).TamLetra
End Property
'AUC 29/11/07
Public Property Get TamLetraPrecio5() As String
    TamLetraPrecio5 = mPosCampo(IND_PRE5).TamLetra
End Property


Public Property Get TamLetraPrecioIVA1() As String
    TamLetraPrecioIVA1 = mPosCampo(IND_PIVA1).TamLetra
End Property

Public Property Get TamLetraPrecioIVA2() As String
    TamLetraPrecioIVA2 = mPosCampo(IND_PIVA2).TamLetra
End Property

Public Property Get TamLetraPrecioIVA3() As String
    TamLetraPrecioIVA3 = mPosCampo(IND_PIVA3).TamLetra
End Property

Public Property Get TamLetraPrecioIVA4() As String
    TamLetraPrecioIVA4 = mPosCampo(IND_PIVA4).TamLetra
End Property
Public Property Get TamLetraPrecioIVA5() As String
    TamLetraPrecioIVA5 = mPosCampo(IND_PIVA5).TamLetra
End Property

Public Property Get TamLetraPrecioIVA6() As String
    TamLetraPrecioIVA6 = mPosCampo(IND_PIVA6).TamLetra
End Property

Public Property Get TamLetraCodigoBarra() As Integer
    TamLetraCodigoBarra = mPosCampo(IND_CBR).TamLetra
End Property

'**********Propiedades para establecer si un campo se imprime ó no************
Public Property Get ImprimirEmpresa() As Boolean
    ImprimirEmpresa = mPosCampo(IND_EMP).Imprimir
End Property

Public Property Get ImprimirCodigo() As Boolean
    ImprimirCodigo = mPosCampo(IND_COD).Imprimir
End Property

Public Property Get ImprimirCodAlterno() As Boolean
    ImprimirCodAlterno = mPosCampo(IND_CA1).Imprimir
End Property

Public Property Get ImprimirDescripcion() As Boolean
    ImprimirDescripcion = mPosCampo(IND_DES).Imprimir
End Property

Public Property Get ImprimirCosto() As Boolean
    ImprimirCosto = mPosCampo(IND_COS).Imprimir
End Property

Public Property Get ImprimirPrecio1() As String
    ImprimirPrecio1 = mPosCampo(IND_PRE1).Imprimir
End Property

Public Property Get ImprimirPrecio2() As String
    ImprimirPrecio2 = mPosCampo(IND_PRE2).Imprimir
End Property

Public Property Get ImprimirPrecio3() As String
    ImprimirPrecio3 = mPosCampo(IND_PRE3).Imprimir
End Property

Public Property Get ImprimirPrecio4() As String
    ImprimirPrecio4 = mPosCampo(IND_PRE4).Imprimir
End Property
Public Property Get ImprimirPrecio5() As String
    ImprimirPrecio5 = mPosCampo(IND_PRE5).Imprimir
End Property
Public Property Get ImprimirPrecio6() As String
    ImprimirPrecio6 = mPosCampo(IND_PRE6).Imprimir
End Property

Public Property Get ImprimirPrecioIVA1() As String
    ImprimirPrecioIVA1 = mPosCampo(IND_PIVA1).Imprimir
End Property

Public Property Get ImprimirPrecioIVA2() As String
    ImprimirPrecioIVA2 = mPosCampo(IND_PIVA2).Imprimir
End Property

Public Property Get ImprimirPrecioIVA3() As String
    ImprimirPrecioIVA3 = mPosCampo(IND_PIVA3).Imprimir
End Property

Public Property Get ImprimirPrecioIVA4() As String
    ImprimirPrecioIVA4 = mPosCampo(IND_PIVA4).Imprimir
End Property
Public Property Get ImprimirPrecioIVA5() As String
    ImprimirPrecioIVA5 = mPosCampo(IND_PIVA5).Imprimir
End Property
Public Property Get ImprimirPrecioIVA6() As String
    ImprimirPrecioIVA6 = mPosCampo(IND_PIVA6).Imprimir
End Property

Public Property Get ImprimirCodigoBarra() As Boolean
    ImprimirCodigoBarra = mPosCampo(IND_CBR).Imprimir
End Property



'**********Propiedades para establecer el numero de caractere a imprimir************
Public Property Get NumLetrasEmpresa() As Integer
    NumLetrasEmpresa = mPosCampo(IND_EMP).Longitud
End Property

Public Property Get NumLetrasCodigo() As Integer
    NumLetrasCodigo = mPosCampo(IND_COD).Longitud
End Property

Public Property Get NumLetrasCodAlterno() As Integer
    NumLetrasCodAlterno = mPosCampo(IND_CA1).Longitud
End Property

Public Property Get NumLetrasDescripcion() As Integer
    NumLetrasDescripcion = mPosCampo(IND_DES).Longitud
End Property

Public Property Get AltoCodigoBarra() As Single
    AltoCodigoBarra = CSng("0" & mPosCampo(IND_CBR).Longitud) * MmTw
End Property

'**********Propiedades para establecer el Formato de los campos************
Public Property Get FormatoEmpresa() As String
    FormatoEmpresa = mPosCampo(IND_EMP).Formato
End Property

Public Property Get FormatoCodigo() As String
    FormatoCodigo = mPosCampo(IND_COD).Formato
End Property

Public Property Get FormatoCodAlterno() As String
    FormatoCodAlterno = mPosCampo(IND_CA1).Formato
End Property

Public Property Get FormatoDescripcion() As String
    FormatoDescripcion = mPosCampo(IND_DES).Formato
End Property

Public Property Get FormatoCosto() As String
    FormatoCosto = mPosCampo(IND_COS).Formato
End Property

Public Property Get FormatoPrecio1() As String
    FormatoPrecio1 = mPosCampo(IND_PRE1).Formato
End Property

Public Property Get FormatoPrecio2() As String
    FormatoPrecio2 = mPosCampo(IND_PRE2).Formato
End Property

Public Property Get FormatoPrecio3() As String
    FormatoPrecio3 = mPosCampo(IND_PRE3).Formato
End Property

Public Property Get FormatoPrecio4() As String
    FormatoPrecio4 = mPosCampo(IND_PRE4).Formato
End Property

Public Property Get FormatoPrecio5() As String
    FormatoPrecio5 = mPosCampo(IND_PRE5).Formato
End Property

Public Property Get FormatoPrecio6() As String
    FormatoPrecio6 = mPosCampo(IND_PRE6).Formato
End Property

Public Property Get FormatoPrecioIVA1() As String
    FormatoPrecioIVA1 = mPosCampo(IND_PIVA1).Formato
End Property

Public Property Get FormatoPrecioIVA2() As String
    FormatoPrecioIVA2 = mPosCampo(IND_PIVA2).Formato
End Property

Public Property Get FormatoPrecioIVA3() As String
    FormatoPrecioIVA3 = mPosCampo(IND_PIVA3).Formato
End Property

Public Property Get FormatoPrecioIVA4() As String
    FormatoPrecioIVA4 = mPosCampo(IND_PIVA4).Formato
End Property
Public Property Get FormatoPrecioIVA5() As String
    FormatoPrecioIVA5 = mPosCampo(IND_PIVA5).Formato
End Property

Public Property Get FormatoPrecioIVA6() As String
    FormatoPrecioIVA6 = mPosCampo(IND_PIVA6).Formato
End Property

Public Property Get FormatoCodigoBarra() As Integer
    FormatoCodigoBarra = mPosCampo(IND_CBR).Formato
End Property



'**********Propiedades para establecer el Prefijo de los campos************
Public Property Get PrefijoEmpresa() As String
    PrefijoEmpresa = mPosCampo(IND_EMP).Prefijo
End Property

Public Property Get PrefijoCodigo() As String
    PrefijoCodigo = mPosCampo(IND_COD).Prefijo
End Property

Public Property Get PrefijoCodAlterno() As String
    PrefijoCodAlterno = mPosCampo(IND_CA1).Prefijo
End Property

Public Property Get PrefijoDescripcion() As String
    PrefijoDescripcion = mPosCampo(IND_DES).Prefijo
End Property

Public Property Get PrefijoCosto() As String
    PrefijoCosto = mPosCampo(IND_COS).Prefijo
End Property

Public Property Get PrefijoPrecio1() As String
    PrefijoPrecio1 = mPosCampo(IND_PRE1).Prefijo
End Property

Public Property Get PrefijoPrecio2() As String
    PrefijoPrecio2 = mPosCampo(IND_PRE2).Prefijo
End Property

Public Property Get PrefijoPrecio3() As String
    PrefijoPrecio3 = mPosCampo(IND_PRE3).Prefijo
End Property

Public Property Get PrefijoPrecio4() As String
    PrefijoPrecio4 = mPosCampo(IND_PRE4).Prefijo
End Property

Public Property Get PrefijoPrecio5() As String
    PrefijoPrecio5 = mPosCampo(IND_PRE5).Prefijo
End Property

Public Property Get PrefijoPrecio6() As String
    PrefijoPrecio6 = mPosCampo(IND_PRE6).Prefijo
End Property
Public Property Get PrefijoPrecioIVA1() As String
    PrefijoPrecioIVA1 = mPosCampo(IND_PIVA1).Prefijo
End Property

Public Property Get PrefijoPrecioIVA2() As String
    PrefijoPrecioIVA2 = mPosCampo(IND_PIVA2).Prefijo
End Property

Public Property Get PrefijoPrecioIVA3() As String
    PrefijoPrecioIVA3 = mPosCampo(IND_PIVA3).Prefijo
End Property

Public Property Get PrefijoPrecioIVA4() As String
    PrefijoPrecioIVA4 = mPosCampo(IND_PIVA4).Prefijo
End Property
Public Property Get PrefijoPrecioIVA5() As String
    PrefijoPrecioIVA5 = mPosCampo(IND_PIVA5).Prefijo
End Property
Public Property Get PrefijoPrecioIVA6() As String
    PrefijoPrecioIVA6 = mPosCampo(IND_PIVA6).Prefijo
End Property
Public Property Get PrefijoCodigoBarra() As String
    PrefijoCodigoBarra = mPosCampo(IND_CBR).Prefijo
End Property
'***Métodos
Public Sub RecuperarConfig()
    Dim cadena As String
    
    'Cadena = GetSetting(APPNAME, App.Title, "MargenesCodBar", "5,10,290,210,20,20,14,14,8,0.02,0")
    
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_MargenesCodBar")) > 0 Then
            cadena = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiPrecio_MargenesCodBar")
        Else
            cadena = "5,10,290,210,20,20,14,14,8,0.02,0,0"
        End If

    
    MargenesCodBar = Split(cadena, ",")
End Sub

Private Sub PosicionesCampos()
    Dim poscampo As String, v As Variant, i As Integer, CONT As Integer
    
    CONT = 0
    
    poscampo = gconfig.PosCampoEMP
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1
    
    poscampo = gconfig.PosCampoCOD
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)

    CONT = CONT + 1
    
    poscampo = gconfig.PosCampoCA1
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1
    
    poscampo = gconfig.PosCampoDES
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1
    
    poscampo = gconfig.PosCampoCOS
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1
    
    poscampo = gconfig.PosCampoPRE1
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1
    
    poscampo = gconfig.PosCampoPRE2
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1
    
    poscampo = gconfig.PosCampoPRE3
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1
    
    poscampo = gconfig.PosCampoPRE4
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1
    
    'AUC  29/11/07
    poscampo = gconfig.PosCampoPRE5
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1
    
    'AUC  precio 6
    poscampo = gconfig.PosCampoPRE6
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1
    
    poscampo = gconfig.PosCampoPIVA1
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1

    poscampo = gconfig.PosCampoPIVA2
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1
    
    poscampo = gconfig.PosCampoPIVA3
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1
    
    poscampo = gconfig.PosCampoPIVA4
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1
    
    'AUC 29/11/07
    poscampo = gconfig.PosCampoPIVA5
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1
    
    poscampo = gconfig.PosCampoPIVA6
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1
    
    poscampo = gconfig.PosCampoCBR
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1
    
    poscampo = gconfig.PosCampoIVG1
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1

    poscampo = gconfig.PosCampoIVGD1
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1
    
    poscampo = gconfig.PosCampoIVG2
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1

    poscampo = gconfig.PosCampoIVGD2
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1

    poscampo = gconfig.PosCampoIVG3
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1
    
    poscampo = gconfig.PosCampoIVGD3
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1

    poscampo = gconfig.PosCampoIVG4
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1
    
    poscampo = gconfig.PosCampoIVGD4
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1

    poscampo = gconfig.PosCampoIVG5
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1
    
    poscampo = gconfig.PosCampoIVGD5
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1
    
    poscampo = gconfig.PosCampoCant
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    CONT = CONT + 1
    
    poscampo = gconfig.PosCampoCodigoDer
    v = Split(poscampo, ";")
    mPosCampo(CONT).X = v(0)
    mPosCampo(CONT).Y = v(1)
    mPosCampo(CONT).Fuente = v(2)
    mPosCampo(CONT).TamLetra = v(3)
    mPosCampo(CONT).Imprimir = v(4)
    mPosCampo(CONT).Longitud = v(5)
    mPosCampo(CONT).Formato = v(6)
    mPosCampo(CONT).Prefijo = v(7)
    
End Sub

Public Sub GrabarConfig()
    Dim i As Integer, cadena As String
    
    'Si se cambió algo se graba
    If mbooModificado Then
        For i = 0 To UBound(MargenesCodBar, 1)
            cadena = cadena & MargenesCodBar(i) & ","
        Next i
        'Quita la última coma
        cadena = Mid$(cadena, 1, Len(cadena) - 1)
        SaveSetting APPNAME, App.Title, "MargenesCodBar", cadena
    End If
End Sub

Public Sub CalcularPosiciones()
    Dim i As Integer, j As Integer, CONT As Integer, num_etiq As Integer
    Dim alto_etiq As Single, ancho_etiq As Single, posx As Single, posy As Single
    
    'Calcula el # total de etiquetas por c/hoja
    num_etiq = Columnas * Filas
    
    'Calcula la dimensión de una sola etiqueta de la hoja
    alto_etiq = (AltoPapel - (MargenSuperior + MargenInferior)) / Filas
    ancho_etiq = (AnchoPapel - (MargenIzquierdo + MargenDerecho)) / Columnas
    
    ReDim Preserve gCodBar(num_etiq)
    CONT = 1
    posx = MargenIzquierdo
    posy = (MargenSuperior - MargenImpresora) + EspacioFilas
    
    'Recupera posiciones de campos a imprimir
    PosicionesCampos
    
    For i = 1 To Filas
        posx = MargenIzquierdo
        For j = 1 To Columnas
            'Posiciones para el nombre de la empresa
            gCodBar(CONT).NombreEmpresa.X = posx + (mPosCampo(IND_EMP).X * MmTw)
            gCodBar(CONT).NombreEmpresa.Y = posy + (mPosCampo(IND_EMP).Y * MmTw)
            
            'Posiciones para el código del item
            gCodBar(CONT).codigoItem.X = posx + (mPosCampo(IND_COD).X * MmTw)
            gCodBar(CONT).codigoItem.Y = posy + (mPosCampo(IND_COD).Y * MmTw)
            
            'Posiciones para el código alterno del item
            gCodBar(CONT).CodAlterno.X = posx + (mPosCampo(IND_CA1).X * MmTw)
            gCodBar(CONT).CodAlterno.Y = posy + (mPosCampo(IND_CA1).Y * MmTw)
                        
            'Posiciones para la descripcion del item.
            gCodBar(CONT).Descripcion.X = posx + (mPosCampo(IND_DES).X * MmTw)
            gCodBar(CONT).Descripcion.Y = posy + (mPosCampo(IND_DES).Y * MmTw)
            
            'Posiciones para el costo
            gCodBar(CONT).costo.X = posx + (mPosCampo(IND_COS).X * MmTw)
            gCodBar(CONT).costo.Y = posy + (mPosCampo(IND_COS).Y * MmTw)
            
            'Posiciones para el precio 1 de venta
            gCodBar(CONT).precio1.X = posx + (mPosCampo(IND_PRE1).X * MmTw)
            gCodBar(CONT).precio1.Y = posy + (mPosCampo(IND_PRE1).Y * MmTw)
            
            'Posiciones para el precio 2 de venta
            gCodBar(CONT).precio2.X = posx + (mPosCampo(IND_PRE2).X * MmTw)
            gCodBar(CONT).precio2.Y = posy + (mPosCampo(IND_PRE2).Y * MmTw)
            
            'Posiciones para el precio 3 de venta
            gCodBar(CONT).precio3.X = posx + (mPosCampo(IND_PRE3).X * MmTw)
            gCodBar(CONT).precio3.Y = posy + (mPosCampo(IND_PRE3).Y * MmTw)

            'Posiciones para el precio 4 de venta
            gCodBar(CONT).precio4.X = posx + (mPosCampo(IND_PRE4).X * MmTw)
            gCodBar(CONT).precio4.Y = posy + (mPosCampo(IND_PRE4).Y * MmTw)

            'Posiciones para el precio 5 de venta
            gCodBar(CONT).precio5.X = posx + (mPosCampo(IND_PRE5).X * MmTw)
            gCodBar(CONT).precio5.Y = posy + (mPosCampo(IND_PRE5).Y * MmTw)
            
            gCodBar(CONT).precio6.X = posx + (mPosCampo(IND_PRE6).X * MmTw)
            gCodBar(CONT).precio6.Y = posy + (mPosCampo(IND_PRE6).Y * MmTw)
    
            'Posiciones para el precioIVA 1 de venta
            gCodBar(CONT).PrecioIVA1.X = posx + (mPosCampo(IND_PIVA1).X * MmTw)
            gCodBar(CONT).PrecioIVA1.Y = posy + (mPosCampo(IND_PIVA1).Y * MmTw)
            
            'Posiciones para el precioIVA 2 de venta
            gCodBar(CONT).PrecioIVA2.X = posx + (mPosCampo(IND_PIVA2).X * MmTw)
            gCodBar(CONT).PrecioIVA2.Y = posy + (mPosCampo(IND_PIVA2).Y * MmTw)
            
            'Posiciones para el precioIVA 3 de venta
            gCodBar(CONT).PrecioIVA3.X = posx + (mPosCampo(IND_PIVA3).X * MmTw)
            gCodBar(CONT).PrecioIVA3.Y = posy + (mPosCampo(IND_PIVA3).Y * MmTw)

            'Posiciones para el precioIVA 4 de venta
            gCodBar(CONT).PrecioIVA4.X = posx + (mPosCampo(IND_PIVA4).X * MmTw)
            gCodBar(CONT).PrecioIVA4.Y = posy + (mPosCampo(IND_PIVA4).Y * MmTw)

            'Posiciones para el precioIVA 5 de venta
            gCodBar(CONT).PrecioIVA5.X = posx + (mPosCampo(IND_PIVA5).X * MmTw)
            gCodBar(CONT).PrecioIVA5.Y = posy + (mPosCampo(IND_PIVA5).Y * MmTw)

            gCodBar(CONT).PrecioIVA6.X = posx + (mPosCampo(IND_PIVA6).X * MmTw)
            gCodBar(CONT).PrecioIVA6.Y = posy + (mPosCampo(IND_PIVA6).Y * MmTw)
            
            'Posiciones para dibujar el código de barras
            gCodBar(CONT).CodigoBarras.X = posx + (mPosCampo(IND_CBR).X * MmTw)
            gCodBar(CONT).CodigoBarras.Y = posy + (mPosCampo(IND_CBR).Y * MmTw)
            
            ''LIRA
            
            gCodBar(CONT).IVGrupo1.X = posx + (mPosCampo(IND_IVG1).X * MmTw)
            gCodBar(CONT).IVGrupo1.Y = posy + (mPosCampo(IND_IVG1).Y * MmTw)
            gCodBar(CONT).IVGrupo1D.X = posx + (mPosCampo(IND_IVGD1).X * MmTw)
            gCodBar(CONT).IVGrupo1D.Y = posy + (mPosCampo(IND_IVGD1).Y * MmTw)
            
            gCodBar(CONT).IVGrupo2.X = posx + (mPosCampo(IND_IVG2).X * MmTw)
            gCodBar(CONT).IVGrupo2.Y = posy + (mPosCampo(IND_IVG2).Y * MmTw)
            gCodBar(CONT).IVGrupo2D.X = posx + (mPosCampo(IND_IVGD2).X * MmTw)
            gCodBar(CONT).IVGrupo2D.Y = posy + (mPosCampo(IND_IVGD2).Y * MmTw)

            gCodBar(CONT).IVGrupo3.X = posx + (mPosCampo(IND_IVG3).X * MmTw)
            gCodBar(CONT).IVGrupo3.Y = posy + (mPosCampo(IND_IVG3).Y * MmTw)
            gCodBar(CONT).IVGrupo3D.X = posx + (mPosCampo(IND_IVGD3).X * MmTw)
            gCodBar(CONT).IVGrupo3D.Y = posy + (mPosCampo(IND_IVGD3).Y * MmTw)
            
            gCodBar(CONT).IVGrupo4.X = posx + (mPosCampo(IND_IVG4).X * MmTw)
            gCodBar(CONT).IVGrupo4.Y = posy + (mPosCampo(IND_IVG4).Y * MmTw)
            gCodBar(CONT).IVGrupo4D.X = posx + (mPosCampo(IND_IVGD4).X * MmTw)
            gCodBar(CONT).IVGrupo4D.Y = posy + (mPosCampo(IND_IVGD4).Y * MmTw)
            
            gCodBar(CONT).IVGrupo5.X = posx + (mPosCampo(IND_IVG5).X * MmTw)
            gCodBar(CONT).IVGrupo5.Y = posy + (mPosCampo(IND_IVG5).Y * MmTw)
            
            gCodBar(CONT).IVGrupo5D.X = posx + (mPosCampo(IND_IVGD5).X * MmTw)
            gCodBar(CONT).IVGrupo5D.Y = posy + (mPosCampo(IND_IVGD5).Y * MmTw)
            
            gCodBar(CONT).Cantidad.X = posx + (mPosCampo(IND_CANT).X * MmTw)
            gCodBar(CONT).Cantidad.Y = posy + (mPosCampo(IND_CANT).Y * MmTw)
            
            gCodBar(CONT).CodigoDer.X = posx + (mPosCampo(IND_CODDER).X * MmTw)
            gCodBar(CONT).CodigoDer.Y = posy + (mPosCampo(IND_CODDER).Y * MmTw)
            
            CONT = CONT + 1
            posx = posx + ancho_etiq
        Next j
        
        posy = posy + alto_etiq + EspacioFilas
    Next i
    
End Sub



Public Property Get FormatoCantidad() As String
    FormatoCantidad = mPosCampo(IND_CANT).Formato
End Property
'--------- fuente
Public Property Get FuenteIVG1() As String
    FuenteIVG1 = mPosCampo(IND_IVG1).Fuente
End Property
Public Property Get FuenteIVGD1() As String
    FuenteIVGD1 = mPosCampo(IND_IVGD1).Fuente
End Property
Public Property Get FuenteIVG2() As String
    FuenteIVG2 = mPosCampo(IND_IVG2).Fuente
End Property
Public Property Get FuenteIVGD2() As String
    FuenteIVGD2 = mPosCampo(IND_IVGD2).Fuente
End Property
Public Property Get FuenteIVG3() As String
    FuenteIVG3 = mPosCampo(IND_IVG3).Fuente
End Property
Public Property Get FuenteIVGD3() As String
    FuenteIVGD3 = mPosCampo(IND_IVGD3).Fuente
End Property
Public Property Get FuenteIVG4() As String
    FuenteIVG4 = mPosCampo(IND_IVG4).Fuente
End Property
Public Property Get FuenteIVGD4() As String
    FuenteIVGD4 = mPosCampo(IND_IVGD4).Fuente
End Property

Public Property Get FuenteIVG5() As String
    FuenteIVG5 = mPosCampo(IND_IVG5).Fuente
End Property
Public Property Get FuenteIVGD5() As String
    FuenteIVGD5 = mPosCampo(IND_IVGD5).Fuente
End Property
Public Property Get FuenteCantidad() As String
    FuenteCantidad = mPosCampo(IND_CANT).Fuente
End Property
Public Property Get FuenteCodigoDer() As String
    FuenteCodigoDer = mPosCampo(IND_CODDER).Fuente
End Property

'--------- Imprimir
Public Property Get ImprimirIVG1() As String
    ImprimirIVG1 = mPosCampo(IND_IVG1).Imprimir
End Property
Public Property Get ImprimirIVGD1() As String
    ImprimirIVGD1 = mPosCampo(IND_IVGD1).Imprimir
End Property
Public Property Get ImprimirIVG2() As String
    ImprimirIVG2 = mPosCampo(IND_IVG2).Imprimir
End Property
Public Property Get ImprimirIVGD2() As String
    ImprimirIVGD2 = mPosCampo(IND_IVGD2).Imprimir
End Property
Public Property Get ImprimirIVG3() As String
    ImprimirIVG3 = mPosCampo(IND_IVG3).Imprimir
End Property
Public Property Get ImprimirIVGD3() As String
    ImprimirIVGD3 = mPosCampo(IND_IVGD3).Imprimir
End Property
Public Property Get ImprimirIVG4() As String
    ImprimirIVG4 = mPosCampo(IND_IVG4).Imprimir
End Property
Public Property Get ImprimirIVGD4() As String
    ImprimirIVGD4 = mPosCampo(IND_IVGD4).Imprimir
End Property
Public Property Get ImprimirIVG5() As String
    ImprimirIVG5 = mPosCampo(IND_IVG5).Imprimir
End Property
Public Property Get ImprimirIVGD5() As String
    ImprimirIVGD5 = mPosCampo(IND_IVGD5).Imprimir
End Property

Public Property Get ImprimirCantidad() As String
    ImprimirCantidad = mPosCampo(IND_CANT).Imprimir
End Property
Public Property Get ImprimirCodigoDer() As String
    ImprimirCodigoDer = mPosCampo(IND_CODDER).Imprimir
End Property

'--------- tamañoletra
Public Property Get TamLetraIVG1() As String
    TamLetraIVG1 = mPosCampo(IND_IVG1).TamLetra
End Property
Public Property Get TamLetraIVGD1() As String
    TamLetraIVGD1 = mPosCampo(IND_IVGD1).TamLetra
End Property
Public Property Get TamLetraIVG2() As String
    TamLetraIVG2 = mPosCampo(IND_IVG2).TamLetra
End Property
Public Property Get TamLetraIVGD2() As String
    TamLetraIVGD2 = mPosCampo(IND_IVGD2).TamLetra
End Property
Public Property Get TamLetraIVG3() As String
    TamLetraIVG3 = mPosCampo(IND_IVG3).TamLetra
End Property
Public Property Get TamLetraIVGD3() As String
    TamLetraIVGD3 = mPosCampo(IND_IVGD3).TamLetra
End Property
Public Property Get TamLetraIVG4() As String
    TamLetraIVG4 = mPosCampo(IND_IVG4).TamLetra
End Property
Public Property Get TamLetraIVGD4() As String
    TamLetraIVGD4 = mPosCampo(IND_IVGD4).TamLetra
End Property

Public Property Get TamLetraIVG5() As String
    TamLetraIVG5 = mPosCampo(IND_IVG5).TamLetra
End Property
Public Property Get TamLetraIVGD5() As String
    TamLetraIVGD5 = mPosCampo(IND_IVGD5).TamLetra
End Property
Public Property Get TamLetraCantidad() As String
    TamLetraCantidad = mPosCampo(IND_CANT).TamLetra
End Property

Public Property Get TamLetraCodigoDer() As String
    TamLetraCodigoDer = mPosCampo(IND_CODDER).TamLetra
End Property
'--------- longitud
Public Property Get LongitudIVG1() As String
    LongitudIVG1 = mPosCampo(IND_IVG1).Longitud
End Property
Public Property Get LongitudIVGD1() As String
    LongitudIVGD1 = mPosCampo(IND_IVGD1).Longitud
End Property
Public Property Get LongitudIVG2() As String
    LongitudIVG2 = mPosCampo(IND_IVG2).Longitud
End Property
Public Property Get LongitudIVGD2() As String
    LongitudIVGD2 = mPosCampo(IND_IVGD2).Longitud
End Property
Public Property Get LongitudIVG3() As String
    LongitudIVG3 = mPosCampo(IND_IVG3).Longitud
End Property
Public Property Get LongitudIVGD3() As String
    LongitudIVGD3 = mPosCampo(IND_IVGD3).Longitud
End Property
Public Property Get LongitudIVG4() As String
    LongitudIVG4 = mPosCampo(IND_IVG4).Longitud
End Property
Public Property Get LongitudIVGD4() As String
    LongitudIVGD4 = mPosCampo(IND_IVGD4).Longitud
End Property

Public Property Get LongitudIVG5() As String
    LongitudIVG5 = mPosCampo(IND_IVG5).Longitud
End Property
Public Property Get LongitudIVGD5() As String
    LongitudIVGD5 = mPosCampo(IND_IVGD5).Longitud
End Property

Public Property Get LongitudCant() As String
    LongitudCant = mPosCampo(IND_CANT).Longitud
End Property

Public Property Get LongitudCodigoDer() As String
    LongitudCodigoDer = mPosCampo(IND_CODDER).Longitud
End Property

Public Property Get PrefijoCantidad() As String
    PrefijoCantidad = mPosCampo(IND_CANT).Prefijo
End Property

