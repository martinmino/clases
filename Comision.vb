Imports System.Data.OracleClient

Public Class Comision

    Private da As OracleDataAdapter
    Private dt As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Dim sql As String

        sql = "SELECT * "
        sql &= "FROM xsalescom "
        sql &= "WHERE tclcod_0 = :tclcod AND comcat_0 = :comcat"

        da = New OracleDataAdapter(sql, cn)
        With da.SelectCommand.Parameters
            .Add("tclcod", OracleType.VarChar)
            .Add("comcat", OracleType.Number)
        End With

    End Sub
    Public Function Abrir(ByVal tclcod As String, ByVal comcat As Integer) As Boolean

        With da.SelectCommand
            .Parameters("tclcod").Value = tclcod
            .Parameters("comcat").Value = comcat
        End With

        dt.Clear()
        da.Fill(dt)

        Return dt.Rows.Count > 0
    End Function

    Public ReadOnly Property Comision(ByVal index As Integer) As Double
        Get
            Dim c As Double = 0
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                c = CDbl(dr("ratrep_" & index.ToString))
            End If

            Return c

        End Get
    End Property

End Class
