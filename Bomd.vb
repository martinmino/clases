Imports System.Data.OracleClient

Public Class Bomd
    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String

        Me.cn = cn

        Sql = "SELECT * FROM bomd WHERE itmref_0=:itmref AND bomseq_0=:bomseq AND bomalt_0=99"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("itmref", OracleType.VarChar)
        da.SelectCommand.Parameters.Add("bomseq", OracleType.VarChar)

    End Sub
    Public Function Abrir(ByVal Art As String, ByVal tipo As String) As Boolean
        With da.SelectCommand
            .Parameters("itmref").Value = Art
            .Parameters("bomseq").Value = tipo
        End With

        dt.Clear()

        Try
            da.Fill(dt)
            Return dt.Rows.Count > 0

        Catch ex As Exception

        End Try

        Return False

    End Function
    Public ReadOnly Property Dias() As Integer
        Get
            Dim dr As DataRow
            dr = dt.Rows(0)

            Return CInt(dr("ydayfreq_0"))
        End Get
    End Property

End Class