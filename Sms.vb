Imports System.Net
Imports System.IO

Public Class Sms
    Const SMS_USR As String = "GEORGIA"
    Const SMS_PWD As String = "geo2838"
    Const SMS_API_ENVIO As String = "http://servicio.smsmasivos.com.ar/enviar_sms.asp?api=1"
    Const SMS_API_OBTENER_ENVIOS As String = "http://servicio.smsmasivos.com.ar/obtener_envios.asp?"

    Private _Texto As String
    Private _Tos As String
    Private _Id As String
    Private _Test As Boolean
    Private _Resp As String = ""

    Public Sub New()
        Nuevo()
    End Sub
    Public Sub Nuevo()
        _Texto = ""
        _Tos = ""
        _Id = ""
        _Test = False
    End Sub
    Public Function Enviar() As Boolean
        Dim Url As String
        Dim Request As WebRequest
        Dim Response As WebResponse
        Dim flg As Boolean = False

        Url = SMS_API_ENVIO
        Url &= "&USUARIO=" & SMS_USR
        Url &= "&CLAVE=" & SMS_PWD
        Url &= "&TEXTO=" & _Texto
        If _Id <> "" Then Url &= "&IDINTERNO=" & _Id
        If _Test Then Url &= "&TEST=1"

        Try
            'Escapo la cadena url
            Url = Uri.EscapeUriString(Url)
            'Valido que estén todos los campos cargados
            If _Texto = "" Then Exit Try
            If _Tos = "" Then Exit Try
            'Agrego el teléfono celular si el número comienza con 15
            'y envio el sms
            Url &= "&TOS=" & _Tos

            'Envio el SMS via HTTP GET
            Request = Net.WebRequest.Create(Url)
            Response = Request.GetResponse

            ' Get the stream containing content returned by the server.
            Dim dataStream As Stream = Response.GetResponseStream()
            ' Open the stream using a StreamReader for easy access.
            Dim reader As New StreamReader(dataStream)
            ' Read the content.
            Dim responseFromServer As String = reader.ReadToEnd()
            ' Display the content.
            _Resp = responseFromServer
            ' Clean up the streams and the response.
            reader.Close()
            Response.Close()

            flg = True

        Catch ex As Exception
        End Try

        Return flg

    End Function
    Public Function Enviar(ByVal Texto As String, ByVal Tos As String, Optional ByVal Id As String = "") As Boolean
        Me.Id = Id
        Me.Texto = Texto
        Me.Tos = Tos
        Return Enviar()
    End Function
    Public Property Texto() As String
        Get
            Return _Texto
        End Get
        Set(ByVal value As String)
            _Texto = value
        End Set
    End Property
    Public Property Tos() As String
        Get
            Return _Tos
        End Get
        Set(ByVal value As String)
            _Tos = value
        End Set
    End Property
    Public Property Id() As String
        Get
            Return _Id
        End Get
        Set(ByVal value As String)
            _Id = value
        End Set
    End Property
    Public Property Test() As Boolean
        Get
            Return _Test
        End Get
        Set(ByVal value As Boolean)
            _Test = value
        End Set
    End Property
    Public ReadOnly Property Respuesta() As String
        Get
            Return _Resp
        End Get
    End Property

    Public ReadOnly Property ObtenerEnviados() As Integer
        Get
            Dim url As String
            Dim i As Integer = 0 'Cantidad SMS enviados
            Dim Request As WebRequest
            Dim Response As WebResponse
            Dim dataStream As Stream
            Dim sr As StreamReader
            Dim txt As String

            url = SMS_API_OBTENER_ENVIOS
            url &= "USUARIO=" & SMS_USR
            url &= "&CLAVE=" & SMS_PWD

            Request = Net.WebRequest.Create(url)
            Response = Request.GetResponse

            dataStream = Response.GetResponseStream()
            sr = New StreamReader(dataStream)
            txt = sr.ReadToEnd()
            If IsNumeric(txt) Then i = CInt(txt)

            Return i

        End Get
    End Property

End Class