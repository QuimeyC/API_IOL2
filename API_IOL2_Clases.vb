Imports API_IOL2.API_IOL2

Partial Public Class API_IOL2

    Public Class EstadoDeCuentas
        Inherits Respuesta
        Public Property Cuentas As List(Of Cuenta)
    End Class

    Public Class Cuenta
        Public Property Numero As String
        Public Property Tipo As String
        Public Property Moneda As String
        Public Property Disponible As Decimal
        Public Property Comprometido As Decimal
        Public Property Saldo As Decimal
        Public Property TitulosValorizados As Decimal
        Public Property Total As Decimal
        Public Property MargenDescubierto As Decimal
        Public Property Saldos As List(Of Saldo)
        Public Property Estado As String


    End Class

    Public Class Saldo
        Public Property Liquidacion As Decimal
        Public Property Saldo As Decimal
        Public Property Comprometido As Decimal
        Public Property Disponible As Decimal
        Public Property DisponibleOperar As Decimal
    End Class



    Public Class Portafolio
        Inherits Respuesta
        Public Property Pais As String
        Public Property Activos As List(Of Activo)

    End Class

    Public Class Activo
        Public Property Cantidad As Decimal
        Public Property Comprometido As Decimal
        Public Property PuntosVariacion As Decimal
        Public Property VariacionDiaria As Decimal
        Public Property UltimoPrecio As Decimal
        Public Property PPC As Decimal
        Public Property GananciaPorcentaje As Decimal
        Public Property GananciaDinero As Decimal
        Public Property Valorizado As Decimal
        Public Property Titulo As Activo_Titulo

    End Class

    Public Class Activo_Titulo
        Public Property Simbolo As String
        Public Property Descripcion As String
        Public Property Pais As String
        Public Property Mercado As String
        Public Property Tipo As String
        Public Property Plazo As String
        Public Property Moneda As String


    End Class



    Public Class Descripcion_Simbolo
        Inherits Respuesta
        Public Property Simbolo As String
        Public Property Descripcion As String
        Public Property Pais As String
        Public Property Mercado As String
        Public Property Tipo As String
        Public Property Plazo As String
        Public Property Moneda As String


    End Class

    Public Class Panel
        Inherits Respuesta
        Public Property Titulos As List(Of Titulo)

    End Class

    Public Class Titulo
        Public Property Simbolo As String
        Public Property Puntas As Punta
        Public Property UltimoPrecio As Decimal
        Public Property VariacionPorcentual As Decimal
        Public Property Apertura As Decimal
        Public Property Maximo As Decimal
        Public Property Minimo As Decimal
        Public Property UltimoCierre As Decimal
        Public Property Volumen As Decimal
        Public Property CantidadOperaciones As Decimal
        Public Property Fecha As String
        Public Property Mercado As String
        Public Property Moneda As String






    End Class


    Public Class Punta
        Public Property CantidadCompra As Decimal
        Public Property PrecioCompra As Decimal
        Public Property PrecioVenta As Decimal
        Public Property CantidadVenta As Decimal
    End Class


    Public Class RespuestaOrden
        Inherits Respuesta
        Public Property Ok As Boolean = True

        Public Property Messages As List(Of Operacion_Message)

        Public Property NumeroOperacion As Integer

        Public Overloads ReadOnly Property TodoOK As Boolean
            Get
                Return Ok
            End Get
        End Property

    End Class

    Public Class Operacion_Message
        Public Property Title As String
        Public Property Description As String
    End Class

    Public Class Ordenes
        Inherits Respuesta
        Public Property Operaciones As RespuestaOperaciones

    End Class


    Public Class RespuestaOperaciones
        Inherits List(Of Orden)
    End Class

    Public Class Orden
        Inherits Respuesta
        Public Property Numero As Integer?
        Public Property FechaOrden As DateTime?
        Public Property Tipo As String
        Public Property Estado As String
        Public Property Mercado As String
        Public Property Simbolo As String
        Public Property Cantidad As Decimal?
        Public Property Monto As Decimal?
        Public Property Modalidad As String
        Public Property Precio As Decimal?
        Public Property FechaOperada As DateTime?
        Public Property CantidadOperada As Decimal?
        Public Property PrecioOperado As Decimal?
        Public Property MontoOperado As Decimal?

    End Class

    '----------------------------------------------
    '/api/v2/{mercado}/Titulos/{simbolo}/Opciones
    '----------------------------------------------
    Public Class Opciones
        Inherits Respuesta

        Public Property ListaDeOpciones As DatosOpciones

        Public Property SimboloSubyacente As String

    End Class
    Public Class DatosOpciones
        Inherits List(Of DatosOpcion)

    End Class

    Public Class DatosOpcion

        Public Property Cotizacion As Cotizacion

        Public Property SimboloSubyacente As String
        Public Property FechaVencimiento As Date
        Public Property TipoOpcion As String
        Public Property Simbolo As String
        Public Property Descripcion As String
        Public Property Pais As String
        Public Property Mercado As String
        Public Property Plazo As String
        Public Property Moneda As String

        Public ReadOnly Property Strike As Decimal
            Get
                Dim s() As String = Split(Descripcion, " ")
                If s.Length > 2 Then
                    Return Val(s(2))
                Else
                    Return 0
                End If
            End Get
        End Property

        Public Overrides Function ToString() As String
            Dim s As String = Simbolo & " " & Descripcion & ",  $" & Cotizacion.UltimoPrecio.ToString("0.000")
            Return s
        End Function

    End Class


    ''' <summary>
    ''' Mantiene los datos la cotización de un activo
    ''' </summary>
    Public Class Cotizacion
        Inherits Respuesta

        Public Property UltimoPrecio As Decimal
        Public Property Variacion As Decimal
        Public Property Apertura As Decimal
        Public Property Maximo As Decimal
        Public Property Minimo As Decimal

        Public Property FechaHora As DateTime
        Public Property Tendencia As String
        Public Property CierreAnterior As Decimal
        Public Property MontoOperado As Decimal
        Public Property VolumenNominal As Decimal
        Public Property PrecioPromeido As Decimal
        Public Property Moneda As String

        Public Property PrecioAjuste As Decimal
        Public Property InteresesAbiertos As Decimal

        Public Property Puntas As List(Of Punta)


        Public Property CantidadOperaciones As Integer



        Overrides Function ToString() As String
            Dim s As String = "Precio: " & Me.UltimoPrecio & "   Variación:" & Me.Variacion & "  Operado:" & Me.MontoOperado & vbNewLine
            s &= Me.Puntas.ToString()
            Return s
        End Function

    End Class
#Region "Clases para manejo de paneles e instrumentos"

    Public Class Pais
        Inherits Respuesta
        Public Property ListaInstrumentos As RespuestaPais

    End Class


    Public Class RespuestaPais
        Inherits List(Of Instrumento)

    End Class
    Public Class Instrumento
        Public Property Instrumento As String
        Public Property Pais As String
    End Class


    Public Class Paneles
        Inherits Respuesta
        Public Property ListaDePaneles As VariosPaneles
    End Class

    Public Class VariosPaneles
        Inherits List(Of UnPanel)
    End Class
    Public Class UnPanel
        Public Property Panel As String
    End Class
#End Region


    '------------------------------------------------------------------------
    ''' <summary>
    ''' Mantiene los datos historicos de un simbolo
    ''' </summary>
    Public Class DatosHistoricos
        Inherits Respuesta

        Public Property Series As List(Of Cotizacion)



        ''' <summary>
        ''' Exportar a formato excel los datos de serie 
        ''' </summary>
        Public Function ExportarCSV(archivo As String, Optional Separador As String = ";") As Boolean
            Dim Buffer As String = "Fecha" & Separador
            Buffer &= "Ultimo" & Separador
            Buffer &= "Àpertura" & Separador
            Buffer &= "Máximo" & Separador
            Buffer &= "Mínimo" & Separador
            Buffer &= "Volumen" & Separador
            Buffer &= "Monto operado" & vbNewLine

            For Each c As Cotizacion In Series
                Buffer &= c.FechaHora.ToString("yyyy-MM-dd") & Separador
                Buffer &= c.UltimoPrecio.ToString("0.00") & Separador
                Buffer &= c.Apertura.ToString("0.00") & Separador
                Buffer &= c.Maximo.ToString("0.00") & Separador
                Buffer &= c.Minimo.ToString("0.00") & Separador
                Buffer &= c.VolumenNominal.ToString("0") & Separador
                Buffer &= c.MontoOperado.ToString("0.00") & vbNewLine
            Next
            Try
                IO.File.WriteAllText(archivo, Buffer)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

    End Class

    ''' <summary>
    ''' Clase Base utilizada para mantener el error y el json de todas las clase que puede devolver la libreria
    ''' </summary>
    Public Class Respuesta
        Public Property DescripcionError As String = ""
        Public Property Json As String = ""

        Public ReadOnly Property TodoOk As Boolean
            Get
                Return DescripcionError.Length = 0
            End Get
        End Property
        Public ReadOnly Property HuboError As Boolean
            Get
                Return Not TodoOk
            End Get
        End Property

    End Class

End Class
