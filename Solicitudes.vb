Imports System.Text
Imports System.Net

Public Class Solicitudes


    ReadOnly Api As API_IOL
    Dim Solicitud As WebRequest


    Public Property EndPoint As String
    Public Property Metodo As String

    Public Property Parametros As Dictionary(Of String, String)
    Public Property Paths As Dictionary(Of String, String)


    ''' <summary>
    ''' Contructor necesita recibir la intancia de la API_IOL
    ''' </summary>
    Public Sub New(api As API_IOL)
        Me.Api = api

        Paths = New Dictionary(Of String, String)
        Parametros = New Dictionary(Of String, String)


    End Sub

    ''' <summary>
    ''' Funcion que se encarga de preparar los parametros, la solicitud y devolver los datos enviado por el servidor
    ''' </summary>
    ''' <param name="PasarAccessToken">Si es falso estamos intentando obtener token o refresh token</param>
    Public Function Send(Optional PasarAccessToken As Boolean = True) As String
        Api.MensajeError = ""
        Try
            If PasarAccessToken Then
                If Not Api.VerificarToken Then
                    Api.MensajeError = "Error: no se pudo obtener token"
                    Return Nothing
                End If
            End If
            Dim Respuesta As String = ""


            'Transforma la lista de parametros en un string con formato formularios web
            Dim sParametros As String = ""
            For Each par As KeyValuePair(Of String, String) In Parametros
                sParametros &= IIf(sParametros.Length > 0, "&", "") & par.Key & "=" & par.Value
            Next

            'Reemplaza los parametros del path del EndPoint con los parametros pasados en Paths
            Dim sEndPoint As String = EndPoint
            For Each par As KeyValuePair(Of String, String) In Paths
                sEndPoint = Strings.Replace(sEndPoint, par.Key, par.Value)
            Next

            'Completa la URL final que se va a enviar
            Dim sURL As String = Api.URL & sEndPoint

            Dim bParametros() As Byte = Nothing

            Select Case Metodo
                Case "GET"
                    'pone los parametros en la url
                    sURL &= "?" & sParametros
                Case "POST", "DELETE"
                    'Convierte los parametros en array de bytes para luego enviarlos como si fuera un formulario web
                    bParametros = Encoding.UTF8.GetBytes(sParametros)
            End Select

            'Inicializa la solicutd
            Solicitud = WebRequest.Create(sURL)
            Solicitud.PreAuthenticate = True
            Solicitud.Method = Metodo

            'Agrga en el encabezado el access token
            If PasarAccessToken Then
                Solicitud.Headers.Add("Authorization", "Bearer " + Api.Token.Access_Token)
            End If

            Select Case Metodo
                Case "GET"

                Case "POST", "DELETE"
                    'Escribir parametros antes de enviar 
                    Solicitud.ContentType = "application/x-www-form-urlencoded"
                    Solicitud.ContentLength = bParametros.Length
                    Using solicitudFlujo = Solicitud.GetRequestStream
                        solicitudFlujo.Write(bParametros, 0, bParametros.Length)
                        solicitudFlujo.Close()
                    End Using

            End Select

            'Obtener los datos del server
            Using respuestaStream = Solicitud.GetResponse.GetResponseStream
                Using lector As New IO.StreamReader(respuestaStream)
                    Respuesta = lector.ReadToEnd()
                End Using
            End Using

            Return Respuesta

        Catch ex As Exception
            Api.MensajeError = ex.Message
            Return ""
        End Try

    End Function
End Class
