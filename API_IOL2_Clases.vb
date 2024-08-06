Partial Public Class API_IOL2

    Public Class EstadoDeCuentas
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
        Public Property Simbolo As String
        Public Property Descripcion As String
        Public Property Pais As String
        Public Property Mercado As String
        Public Property Tipo As String
        Public Property Plazo As String
        Public Property Moneda As String


    End Class

    Public Class Panel
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

End Class
