Imports System.Data.OracleClient

Public Class Correo
    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String

        Me.cn = cn

        Sql = "SELECT * FROM xcorreo WHERE num_0 = :num_0 ORDER BY dat_0"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("num_0", OracleType.VarChar)

    End Sub

    Public Function Abrir(ByVal Num As String) As Boolean
        dt.Clear()

        da.SelectCommand.Parameters("num_0").Value = Num
        da.Fill(dt)

        Return dt.Rows.Count > 0

    End Function

    Public ReadOnly Property Fecha(ByVal i As Integer) As Date
        Get
            Dim dr As DataRow = dt.Rows(i)
            Return CDate(dr("dat_0"))
        End Get
    End Property
    Public ReadOnly Property Count() As Integer
        Get
            Return dt.Rows.Count
        End Get
    End Property
    Public ReadOnly Property UltimaFecha() As Date
        Get
            Dim dr As DataRow
            dr = dt.Rows(dt.Rows.Count - 1)
            Return CDate(dr("dat_0"))
        End Get
    End Property

End Class
