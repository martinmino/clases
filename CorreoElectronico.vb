Imports System.Net.Mail
Imports System.Text.RegularExpressions
Imports System.IO

Public Class CorreoElectronico
    Implements IDisposable

    Private smtp As New SmtpClient
    Private eMail As MailMessage

    Public Sub New()

        With smtp
            .Host = "smtp-relay.gmail.com"
            .EnableSsl = False
            .UseDefaultCredentials = False
            .Port = 587
        End With

        Nuevo()

    End Sub
    Public Sub Nuevo()
        If eMail IsNot Nothing Then eMail.Dispose()
        eMail = New MailMessage
    End Sub
    Public Sub Remitente(ByVal eMail As String, Optional ByVal Nombre As String = "Grupo Georgia - Prevención y extinción de incendios")
        Try
            If Me.ValidarMail(eMail) Then
                Me.eMail.From = New MailAddress(eMail, Nombre)
            End If

        Catch ex As Exception

        End Try
    End Sub
    Public Sub AgregarDestinatario(ByVal eMail As String, Optional ByVal Oculto As Boolean = False)
        If Oculto Then
            CargarMails(eMail, 2)
        Else
            CargarMails(eMail, 0)
        End If

    End Sub
    Public Sub ResponderA(ByVal eMail As String)
        If ValidarMail(eMail) Then Me.eMail.ReplyTo = New MailAddress(eMail)
    End Sub
    Public Sub AgregarDestinatarioCopia(ByVal eMail As String)
        CargarMails(eMail, 1)
    End Sub
    Public Sub AgregarDestinatarioArchivo(ByVal Arch As String, Optional ByVal Tipo As Integer = 0)
        Dim sr As StreamReader
        Dim Mail As String

        sr = New StreamReader(Arch)

        Do
            Mail = sr.ReadLine()

            If Not Mail Is Nothing Then CargarMails(Mail, Tipo)

        Loop Until Mail Is Nothing

        sr.Close()

    End Sub
    Public Function AdjuntarArchivo(ByVal Archivo As String) As Boolean
        Dim flg As Boolean = False

        Try
            If File.Exists(Archivo) Then
                eMail.Attachments.Add(New Attachment(Archivo))
                flg = True
            End If

        Catch ex As Exception

        End Try

        Return True

    End Function
    Public Function ExisteMail(ByVal Mail As String) As Boolean
        For i = 0 To eMail.To.Count - 1
            If eMail.To(i).Address = Mail Then Return True
        Next
        For i = 0 To eMail.CC.Count - 1
            If eMail.CC(i).Address = Mail Then Return True
        Next
        For i = 0 To eMail.bcc.Count - 1
            If eMail.Bcc(i).Address = Mail Then Return True
        Next

        Return False
    End Function
    Public Function CuerpoDesdeArchivo(ByVal Archivo As String) As Boolean
        Dim flg As Boolean = False
        Dim sr As StreamReader

        Try
            If File.Exists(Archivo) Then
                sr = New StreamReader(Archivo)
                Me.Cuerpo = sr.ReadToEnd
                sr.Close()
                sr.Dispose()
                flg = True
            End If

        Catch ex As Exception

        End Try

        Return flg

    End Function
    Public Sub Enviar(Optional ByVal Archivar As Boolean = True)

        With eMail
            If .To.Count > 0 Or .CC.Count > 0 Or .Bcc.Count > 0 Then
                If Archivar Then .Bcc.Add("no-responder@georgia.com.ar")
                smtp.Send(eMail)
            End If

        End With
    End Sub

    Private Sub CargarMails(ByVal Mail As String, ByVal Tipo As Integer)
        Dim Mails() = Split(Mail, ";")
        Dim Existe As Boolean
        Dim i As Integer

        For Each Mail In Mails
            Existe = False
            Mail = Mail.Trim.ToLower

            If ValidarMail(Mail) Then
                For i = 0 To eMail.To.Count - 1
                    If eMail.To(i).Address = Mail Then Existe = True
                Next
                For i = 0 To eMail.CC.Count - 1
                    If eMail.CC(i).Address = Mail Then Existe = True
                Next
                For i = 0 To eMail.Bcc.Count - 1
                    If eMail.Bcc(i).Address = Mail Then Existe = True
                Next

                Try
                    If Not Existe Then
                        If Tipo = 0 Then eMail.To.Add(Mail)
                        If Tipo = 1 Then eMail.CC.Add(Mail)
                        If Tipo = 2 Then eMail.Bcc.Add(Mail)
                    End If

                Catch ex As Exception

                End Try

            End If
        Next

    End Sub
    Public Function ValidarMail(ByVal sMail As String) As Boolean
        Dim Mails() = Split(sMail, ";")

        For Each Mail As String In Mails
            Mail = Mail.Trim

            If Regex.IsMatch(Mail, "^([\w-]+\.)*?[\w-]+@[\w-]+\.([\w-]+\.)*?[\w]+$") Then
                Return True
            End If
        Next

        Return False

    End Function

    Public Property Asunto() As String
        Get
            Return eMail.Subject
        End Get
        Set(ByVal value As String)
            eMail.Subject = value
        End Set
    End Property
    Public Property Cuerpo() As String
        Get
            Return eMail.Body
        End Get
        Set(ByVal value As String)
            eMail.Body = value
        End Set
    End Property
    Public Property EsHtml() As Boolean
        Get
            Return eMail.IsBodyHtml
        End Get
        Set(ByVal value As Boolean)
            eMail.IsBodyHtml = value
        End Set
    End Property
    Public ReadOnly Property CantidadTo() As Integer
        Get
            Return eMail.To.Count
        End Get
    End Property
    Public ReadOnly Property CantidadCC() As Integer
        Get
            Return eMail.CC.Count
        End Get
    End Property
    Public Property Prioridad() As System.Net.Mail.MailPriority
        Get
            Return eMail.Priority
        End Get
        Set(ByVal value As System.Net.Mail.MailPriority)
            eMail.Priority = value
        End Set
    End Property
    Public ReadOnly Property MailObject() As MailMessage
        Get
            Return eMail
        End Get
    End Property
    Public ReadOnly Property Para() As MailAddressCollection
        Get
            Return eMail.To
        End Get
    End Property
    Private disposedValue As Boolean = False        ' Para detectar llamadas redundantes

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: Liberar otro estado (objetos administrados).

                Try
                    eMail.Dispose()

                Catch ex As Exception

                End Try

            End If

            ' TODO: Liberar su propio estado (objetos no administrados).
            ' TODO: Establecer campos grandes como Null.
        End If
        Me.disposedValue = True

    End Sub

#Region " IDisposable Support "
    ' Visual Basic agregó este código para implementar correctamente el modelo descartable.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' No cambie este código. Coloque el código de limpieza en Dispose (ByVal que se dispone como Boolean).
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class