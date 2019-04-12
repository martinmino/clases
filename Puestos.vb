Imports System.Data.OracleClient
Imports System.ComponentModel

Public Class Puestos
    Inherits BindingList(Of Puesto)

    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String

        Me.cn = cn

        Sql = "SELECT * FROM xpuestos WHERE sector_0 = :sector_0"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("sector", OracleType.Number)
    End Sub
    Public Sub AbrirPuestos(ByVal Sector As Long)
        dt.Clear()
        Me.Clear()

        da.SelectCommand.Parameters("sector").Value = Sector
        da.Fill(dt)

        For Each dr As DataRow In dt.Rows
            Dim p As New Puesto(cn)
            If p.Abrir(CLng(dr(0))) Then Me.Add(p)
        Next

    End Sub


End Class
