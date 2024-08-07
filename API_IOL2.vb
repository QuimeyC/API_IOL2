Imports System.Net
Imports System.Reflection.Emit
Imports System.Text
Imports API_IOL2.API_IOL2
Imports Newtonsoft


Public Class API_IOL

    Private Usuario As String
    Private Clave As String
    Private ReadOnly CarpetaInterna As String

    Public Token As New Token

    Public mensajeError As String

    Public Property URL = "https://api.invertironline.com/"

    Public Sub New(CarpetaInterna As String)
        Me.CarpetaInterna = CarpetaInterna

    End Sub

    Public Function Login(Usuario As String, Clave As String) As Boolean
        Me.Usuario = Usuario
        Me.Clave = Clave

        mensajeError = ""
        Try

            Dim sData As String = "grant_type=password&username=" & Me.Usuario & "&password=" & Me.Clave
            Dim data = Encoding.UTF8.GetBytes(sData)
            Dim respuesta As String = ""
            Dim solicitud As WebRequest = WebRequest.Create(Me.URL + "token")
            solicitud.Method = "POST"
            solicitud.ContentType = "application/x-www-form-urlencoded"
            solicitud.ContentLength = data.Length

            Using solicitudStream = solicitud.GetRequestStream()
                solicitudStream.Write(data, 0, data.Length)
                solicitudStream.Close()
            End Using

            Using respuestaStream = solicitud.GetResponse.GetResponseStream
                Using reader As New IO.StreamReader(respuestaStream)
                    respuesta = reader.ReadToEnd()
                End Using
            End Using

            Token = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Token)(respuesta)

            If Token.Access_Token.Length > 0 And Token.Refresh_Token.Length > 0 Then
                Token.Time_Token = DateAdd(DateInterval.Minute, 15, Now)
                Return True
            Else
                mensajeError = "No se pudo autenticar"
                Return False
            End If



        Catch ex As Exception
            mensajeError = ex.Message
            Return False
        End Try

    End Function
    Private Function EstamosEnRangoTiempo() As Boolean
        Return DateDiff(DateInterval.Second, Now, Me.Token.Time_Token) <= 0
    End Function


    Private Function VerificarToken() As Boolean
        If EstamosEnRangoTiempo() Then
            If Not RefrescarToken() Then
                If Not Login(Me.Usuario, Me.Clave) Then
                    Return False
                End If
            End If
        End If
        Return True
    End Function
    Public Function ObtenerEstadoDeCuentasOLD() As EstadoDeCuentas

        mensajeError = ""
        Try
            If Not VerificarToken() Then
                mensajeError = "No se pudo obtener el token"
                Return Nothing
            End If
            Dim respuesta As String = ""
            Dim solicitud As WebRequest = WebRequest.Create(Me.URL + "api/v2/estadocuenta")
            solicitud.Method = "GET"
            solicitud.Headers.Add("Authorization", "Bearer " + Me.Token.Access_Token)
            solicitud.PreAuthenticate = True

            Using respuestaStream = solicitud.GetResponse.GetResponseStream
                Using reader As New IO.StreamReader(respuestaStream)
                    respuesta = reader.ReadToEnd()
                End Using
            End Using

            Dim Estado As EstadoDeCuentas = Newtonsoft.Json.JsonConvert.DeserializeObject(Of EstadoDeCuentas)(respuesta)


            If Estado IsNot Nothing Then
                Return Estado
            Else
                mensajeError = "Error: No se pudo obtener el estado de cuenta"
                Return Nothing
            End If

        Catch ex As Exception
            mensajeError = ex.Message
            Return Nothing
        End Try
    End Function

    Public Function ObtenerEstadoDeCuentas() As EstadoDeCuentas
        Dim Solicitud As New Solicitudes(Me) With {
            .EndPoint = "api/v2/estadocuenta",
            .Metodo = "GET"
        }

        Dim Respuesta As String = Solicitud.Send

        Dim Resp As EstadoDeCuentas = JSonToObject(Of EstadoDeCuentas)(Respuesta)
        If Resp Is Nothing Then
            Return New EstadoDeCuentas() With {
                .Json = Respuesta,
                .DescripcionError = "Error: No se pudo obtener estdo de cuenta" & vbNewLine & mensajeError
            }
        Else
            Return Resp
        End If

    End Function

    Public Function RefrescarToken() As Boolean
        Me.Usuario = Usuario
        Me.Clave = Clave

        mensajeError = ""
        Try

            Dim sData As String = "grant_type=refresh_token&refresh_token=" + Me.Token.Refresh_Token
            Dim data = Encoding.UTF8.GetBytes(sData)
            Dim respuesta As String = ""
            Dim solicitud As WebRequest = WebRequest.Create(Me.URL + "token")
            solicitud.Method = "POST"
            solicitud.ContentType = "application/x-www-form-urlencoded"
            solicitud.ContentLength = data.Length

            Using solicitudStream = solicitud.GetRequestStream()
                solicitudStream.Write(data, 0, data.Length)
                solicitudStream.Close()
            End Using

            Using respuestaStream = solicitud.GetResponse.GetResponseStream
                Using reader As New IO.StreamReader(respuestaStream)
                    respuesta = reader.ReadToEnd()
                End Using
            End Using

            Me.Token = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Token)(respuesta)

            If Token.Access_Token.Length > 0 And Token.Refresh_Token.Length > 0 Then
                Token.Time_Token = DateAdd(DateInterval.Minute, 15, Now)
                Return True
            Else
                mensajeError = "No se pudo obtener el token"
                Return False
            End If

        Catch ex As Exception
            mensajeError = ex.Message
            Return False
        End Try

    End Function

    Public Function ObtenerPortafolio(Optional Pais As String = "argentina") As Portafolio

        mensajeError = ""
        Try
            If Not VerificarToken() Then
                mensajeError = "No se pudo obtener el token"
                Return Nothing
            End If
            Dim respuesta As String = ""
            Dim solicitud As WebRequest = WebRequest.Create(Me.URL + "api/v2/portafolio/" + Pais)
            solicitud.Method = "GET"
            solicitud.Headers.Add("Authorization", "Bearer " + Me.Token.Access_Token)
            solicitud.PreAuthenticate = True

            Using respuestaStream = solicitud.GetResponse.GetResponseStream
                Using reader As New IO.StreamReader(respuestaStream)
                    respuesta = reader.ReadToEnd()
                End Using
            End Using

            Dim portafolio As Portafolio = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Portafolio)(respuesta)


            If portafolio IsNot Nothing Then
                Return portafolio
            Else
                mensajeError = "Error: No se pudo obtener el portafolio"
                Return Nothing
            End If

        Catch ex As Exception
            mensajeError = ex.Message
            Return Nothing
        End Try
    End Function

    Public Function ObtenerDescripcion(Simbolo As String, Optional Mercado As String = "bcba") As Descripcion_Simbolo

        mensajeError = ""
        Try
            If Not VerificarToken() Then
                mensajeError = "No se pudo obtener el token"
                Return Nothing
            End If
            Dim respuesta As String = ""
            Dim solicitud As WebRequest = WebRequest.Create(Me.URL + "api/v2/" & Mercado & "/Titulos/" & Simbolo)
            solicitud.Method = "GET"
            solicitud.Headers.Add("Authorization", "Bearer " + Me.Token.Access_Token)
            solicitud.PreAuthenticate = True

            Using respuestaStream = solicitud.GetResponse.GetResponseStream
                Using reader As New IO.StreamReader(respuestaStream)
                    respuesta = reader.ReadToEnd()
                End Using
            End Using

            Dim descripcion As Descripcion_Simbolo = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Descripcion_Simbolo)(respuesta)


            If descripcion IsNot Nothing Then
                Return descripcion
            Else
                mensajeError = "Error: No se pudo obtener el dato del simbolo solicitado"
                Return Nothing
            End If

        Catch ex As Exception
            mensajeError = ex.Message
            Return Nothing
        End Try
    End Function





    Public Function ObtenerCotizacionesPanel(Panel As String, Instrumento As String, Optional Pais As String = "argentina") As Panel

        mensajeError = ""
        Try
            If Not VerificarToken() Then
                mensajeError = "No se pudo obtener el token"
                Return Nothing
            End If
            Dim Parametros As String = "panelCotizacion.instrumento" & Instrumento & "panelCotizacion.panel" & Panel
            Parametros &= "panelCotizacion.pais=" & Pais
            Dim respuesta As String = ""
            Dim solicitud As WebRequest = WebRequest.Create(Me.URL + "api/v2/Cotizaciones/" & Instrumento & "/" & Panel & "/" & Pais & "?" & Parametros)
            solicitud.Method = "GET"
            solicitud.Headers.Add("Authorization", "Bearer " + Me.Token.Access_Token)
            solicitud.PreAuthenticate = True

            Using respuestaStream = solicitud.GetResponse.GetResponseStream
                Using reader As New IO.StreamReader(respuestaStream)
                    respuesta = reader.ReadToEnd()
                End Using
            End Using

            Dim p As Panel = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Panel)(respuesta)


            If p IsNot Nothing Then
                Return p
            Else
                mensajeError = "Error: No se pudieron obtener los datos del panel "
                Return Nothing
            End If

        Catch ex As Exception
            mensajeError = ex.Message
            Return Nothing
        End Try
    End Function

    Public Function ObtenerOpciones(Simbolo As String, Optional Mercado As String = "bcba") As Opciones

        Dim Solicitud As New Solicitudes(Me) With {
            .EndPoint = "/api/v2/{mercado}/Titulos/{simbolo}/Opciones",
            .Metodo = "GET"
        }
        Solicitud.Paths.Add("{simbolo}", Simbolo)
        Solicitud.Paths.Add("{mercado}", Mercado)

        Dim Respuesta As String = Solicitud.Send
        Dim Resp As DatosOpciones = Newtonsoft.Json.JsonConvert.DeserializeObject(Of DatosOpciones)(Respuesta)
        If Resp Is Nothing Then
            Return New Opciones() With {
            .DescripcionError = mensajeError,
            .Json = Respuesta,
            .ListaDeOpciones = New List(Of DatosOpcion)
}
        Else
            Return New Opciones With {
            .SimboloSubyacente = Simbolo,
            .ListaDeOpciones = Resp
            }
        End If


    End Function

    Public Function ObtenerCotizacion(Simbolo As String, Optional plazo As String = "T2", Optional Mercado As String = "BCBA") As Cotizacion
        Dim Solicitud As New Solicitudes(Me) With {
            .EndPoint = "api/v2/{Mercado}/Titulos/{Simbolo}/Cotizacion",
            .Metodo = "GET"}
        Solicitud.Paths.Add("{Mercado}", Mercado)
        Solicitud.Paths.Add("{Simbolo}", Simbolo)

        Solicitud.Parametros.Add("model.simbolo", Simbolo)
        Solicitud.Parametros.Add("model.mercado", Mercado)
        Solicitud.Parametros.Add("model.plazo", plazo)

        Dim Respuesta As String = Solicitud.Send
        Dim Resp As Cotizacion = JSonToObject(Of Cotizacion)(Respuesta)
        If Resp Is Nothing Then
            Return New Cotizacion() With {
                .DescripcionError = mensajeError,
                .Json = Respuesta
            }
        Else
            Return Resp
        End If


    End Function

    Public Function ObtenerInstrumentos(Optional pais As String = "argentina") As Pais
        Dim Solicitud As New Solicitudes(Me) With {
            .EndPoint = "api/v2/{pais}/Titulos/Cotizacion/Instrumentos",
            .Metodo = "GET"
        }
        Solicitud.Paths.Add("{pais}", pais)

        Dim Respuesta As String = Solicitud.Send

        Dim Resp As RespuestaPais = Newtonsoft.Json.JsonConvert.DeserializeObject(Of RespuestaPais)(Respuesta)
        If Resp Is Nothing Then
            Return New Pais() With {
                .DescripcionError = mensajeError,
                .Json = Respuesta,
                .ListaInstrumentos = New RespuestaPais}
        Else
            Return New Pais() With {
                .DescripcionError = "",
                .Json = Respuesta,
                .ListaInstrumentos = Resp}
        End If
    End Function
    Public Function ObtenerNombrePaneles(Instrumento As String, Optional Pais As String = "argentina") As Paneles
        Dim Solicitud As New Solicitudes(Me) With {
            .EndPoint = "api/v2/{pais}/Titulos/Cotizacion/Paneles/{instrumento}",
            .Metodo = "GET"
        }
        Solicitud.Paths.Add("{pais}", Pais)
        Solicitud.Paths.Add("{instrumento}", Instrumento)

        Dim Respuesta As String = Solicitud.Send

        Dim Resp As VariosPaneles = Newtonsoft.Json.JsonConvert.DeserializeObject(Of VariosPaneles)(Respuesta)
        If Resp Is Nothing Then
            Return New Paneles() With {
                .DescripcionError = mensajeError,
                .Json = Respuesta,
                .ListaDePaneles = New VariosPaneles}
        Else
            Return New Paneles() With {
                .DescripcionError = "",
                .Json = Respuesta,
                .ListaDePaneles = Resp}
        End If
    End Function

    Public Function Vender(Simbolo As String, Cantidad As Integer, Precio As Decimal, Optional Plazo As String = "t2", Optional DiazValidez As Integer = 0, Optional Mercado As String = "BCBA") As RespuestaOrden
        Return Operar("Vender", Simbolo, Cantidad, Precio, Plazo, DiazValidez, Mercado)

    End Function

    Public Function Comprar(Simbolo As String, Cantidad As Integer, Precio As Decimal, Optional Plazo As String = "t2", Optional DiazValidez As Integer = 0, Optional Mercado As String = "BCBA") As RespuestaOrden
        Return Operar("Comprar", Simbolo, Cantidad, Precio, Plazo, DiazValidez, Mercado)

    End Function

    Private Function Operar(Sentido As String, Simbolo As String, Cantidad As Integer, Precio As Decimal, Optional plazo As String = "t2", Optional DiasValidez As Integer = 0, Optional Mercado As String = "BCBA") As RespuestaOrden

        Dim Solicitud As New Solicitudes(Me) With {
            .EndPoint = "api/v2/operar/" & Sentido,
            .Metodo = "POST"
        }

        Dim FechaVto As DateTime = DateAdd(DateInterval.Day, DiasValidez, Now)
        Dim sFechaVto As String = FechaVto.Year & "-" & FechaVto.Month & "-" & FechaVto.Day & "T17:59:59"

        With Solicitud
            .Parametros.Add("Mercado", Mercado)
            .Parametros.Add("Validez", sFechaVto)
            .Parametros.Add("simbolo", Simbolo)
            .Parametros.Add("cantidad", Cantidad)
            .Parametros.Add("precio", Precio.ToString("0.00").Replace(",", "."))
            .Parametros.Add("plazo", plazo)

        End With

        Dim Respuesta As String = Solicitud.Send

        Dim Resp As RespuestaOrden = JSonToObject(Of RespuestaOrden)(Respuesta)
        If Resp Is Nothing Then
            Return New RespuestaOrden With {
                .Json = Respuesta,
                .DescripcionError = mensajeError}
        Else
            Return Resp
        End If

    End Function



    Public Function RefreshToken() As Boolean
        mensajeError = ""
        Return False
End Function



    Public Class Token
    Public Property Access_Token As String
    Public Property Refresh_Token As String
    Public Property Time_Token As DateTime

End Class

#Region "Funciones Auxiliares"

Public Function JSonToObject(Of T)(JSon As String) As T
    Try
        If JSon = "" Then
            Return Nothing
        End If
        Dim r As Object = Newtonsoft.Json.JsonConvert.DeserializeObject(Of T)(JSon)
        CType(r, Respuesta).DescripcionError = ""
        CType(r, Respuesta).Json = JSon
        Return r
    Catch ex As Exception
        MensajeError = ex.Message
        Return Nothing
    End Try
End Function

#End Region
Public Sub GuardarEnAchivo(archivo As String)
    Try

        IO.File.WriteAllText(archivo, Json.JsonConvert.SerializeObject(Me, Json.Formatting.Indented))
    Catch ex As Exception

    End Try
End Sub
Public Shared Function CargarDesdeArchivo(archivo As String) As Token
    Try
        Dim sjson As String = IO.File.ReadAllText(archivo)
        Dim token As Token = Json.JsonConvert.DeserializeObject(Of Token)(sjson)
        Return token
    Catch ex As Exception
        Return New Token
    End Try
End Function
End Class

