Imports System.Data.OracleClient

Public Class ComisionAgua
    Private cn As OracleConnection
    Private dt As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Dim da As OracleDataAdapter
        Dim sql As String

        Me.cn = cn

        sql = "SELECT * FROM xcomagua ORDER BY base_0 DESC"
        da = New OracleDataAdapter(sql, cn)
        da.Fill(dt)
        da.Dispose()
    End Sub
    Public Function ObtenerAlicuota(ByVal Importe As Double) As Double
        Dim Alicuota As Double = 0

        For Each dr As DataRow In dt.Rows
            If Importe >= CDbl(dr("base_0")) Then
                Alicuota = CDbl(dr("alicuota_0"))
                Exit For
            End If
        Next

        Return Alicuota

    End Function

End Class
