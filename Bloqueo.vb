Imports System.Data.OracleClient

Public Class Bloqueo
    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn

        Dim Sql As String

        Sql = "SELECT * "
        Sql &= "FROM apllck "
        Sql &= "WHERE lcksym_0 = :p"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("p", OracleType.VarChar)

    End Sub
    Public Function EstaBloqueado(ByVal Objeto As String, ByVal Numero As String) As Boolean
        da.SelectCommand.Parameters("p").Value = Objeto.ToUpper & Numero.ToUpper

        dt.Clear()
        da.Fill(dt)

        Return dt.Rows.Count > 0
    End Function

End Class
