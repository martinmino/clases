Imports System.Data.OracleClient

Public Class itmmvt

    Shared Sub Sumar(ByVal cn As OracleConnection, ByVal Articulo As String, ByVal Planta As String, ByVal cant As Long)
        Dim da As OracleDataAdapter
        Dim Sql As String
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim i As Long = 0

        Sql = "SELECT * FROM itmmvt WHERE itmref_0 = :itmref AND stofcy_0 = :stofcy"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("itmref", OracleType.VarChar).Value = Articulo
        da.SelectCommand.Parameters.Add("stofcy", OracleType.VarChar).Value = Planta

        da.UpdateCommand = New OracleCommandBuilder(da).GetUpdateCommand

        Try
            da.Fill(dt)

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)

                i = CLng(dr("salsto_0"))
                i += cant

                dr.BeginEdit()
                dr("salsto_0") = i
                dr.EndEdit()

                da.Update(dt)

            End If

        Catch ex As Exception

        Finally
            da.Dispose()

        End Try

    End Sub

End Class