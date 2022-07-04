Imports System.Text.RegularExpressions
Imports System.Data.OracleClient

Public Class Utiles

    Shared Function ValidarTelefono(ByVal Nro As String) As Boolean
        Dim n As Integer = 0
        Dim flg As Boolean = True

        Nro = Nro.Trim
        If Nro.Length > 0 Then
            For i = 0 To Nro.Length - 1
                If IsNumeric(Nro) Then
                    n += 1
                Else
                    flg = False
                End If
                If Not flg Then Exit For
            Next
            'El numero no debe comenzar con 0
            If Nro.Substring(0, 1) = "0" Then flg = False
        End If

        Return flg And n = 10
    End Function
    Shared Function ValidarMails(ByVal Mails As String) As Boolean
        Dim MailsArray() As String
        Dim flg As Boolean = True

        'Separo los mails por ;
        MailsArray = Split(Mails, ";")

        'Recorro todos los mails del array
        For i As Integer = 0 To UBound(MailsArray)
            If Not ValidarMail(MailsArray(i)) Then flg = False
            If Not flg Then Exit For
        Next

        Return flg
    End Function
    Shared Function ValidarMail(ByVal Mail As String) As Boolean
        Return Regex.IsMatch(Mail.Trim, "^([\w-]+\.)*?[\w-]+@[\w-]+\.([\w-]+\.)*?[\w]+$")
    End Function
    Shared Function ValidarCP(ByVal Codigo As String) As Boolean
        Dim i As Integer
        Dim flg As Boolean = True

        Codigo = Codigo.Trim

        For i = 0 To Codigo.Length - 1
            If Not IsNumeric(Codigo.Substring(i, 1)) Then flg = False
        Next

        Return flg And Codigo.Length = 4

    End Function
    Shared Function ValidarHora(ByVal hora As String) As Boolean
        Dim c As Char

        'Si cadena tiene menos de 4 caracteres la hora es invalida
        If hora.Trim.Length < 4 Then Return False

        'Valido que todos los caracteres sean numeros
        For i As Integer = 0 To 3
            c = hora.Chars(i)
            If Not IsNumeric(c) Then Return False
        Next

        Dim hh As Integer = CInt(hora.Substring(0, 2))
        Dim mm As Integer = CInt(hora.Substring(2, 2))

        If hh > 23 Then Return False
        If mm > 59 Then Return False

        Return True

    End Function
    Shared Function ValidarFranjasHorarias(ByVal Desde1 As String, ByVal Hasta1 As String, ByVal Desde2 As String, ByVal Hasta2 As String, ByRef e As String) As Boolean
        e = ""

        If Not ValidarHora(Desde1) Then e = "Error formato horario"
        If Not ValidarHora(Hasta1) Then e = "Error formato horario"
        If Not ValidarHora(Desde2) Then e = "Error formato horario"
        If Not ValidarHora(Hasta2) Then e = "Error formato horario"

        'Validacion de primer franja horaria
        If e = "" Then 'And Desde1 <> "0000" Or Hasta1 <> "0000" Then
            If Desde1 >= Hasta1 Then
                e = "Error en primer franja horaria"
            End If
        End If

        'Validacion de la segunda franja horaria
        If e = "" AndAlso Desde2 <> "0000" Or Hasta2 <> "0000" Then
            If Hasta1 > Desde2 Then
                e = "Error en la segunda franja horaria"
            End If
            If Desde2 > Hasta2 Then
                e = "Error en la segunda franja horaria"
            End If
        End If

        Return (e = "")

    End Function
    Shared Function ValorParametro(ByVal cn As OracleConnection, ByVal Capitulo As String, ByVal Parametro As String) As Integer
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim i As Integer
        Dim dt As New DataTable

        Sql = "select val.valeur_0 "
        Sql &= "from adopar par inner join "
        Sql &= "     adoval val on (par.param_0 = val.param_0) "
        Sql &= "where par.chapitre_0 = :capitulo and "
        Sql &= "      par.param_0 = :parametro"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("capitulo", OracleType.VarChar).Value = Capitulo
        da.SelectCommand.Parameters.Add("parametro", OracleType.VarChar).Value = Parametro
        da.Fill(dt)

        i = CInt(dt.Rows(0).Item(0))

        dt.Dispose() : dt = Nothing
        da.Dispose() : dt = Nothing

        Return i

    End Function
End Class