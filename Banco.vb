Imports System.Data.OracleClient

Public Class Banco
    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable
    Private dr As DataRow

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String = ""
        Me.cn = cn

        Sql = "SELECT * from bank WHERE ban_0 = :ban"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("ban", OracleType.VarChar)

    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal CodigoBanco As String)
        Me.New(cn)
        Abrir(CodigoBanco)
    End Sub
    Public Sub Abrir(ByVal CodigoBanco As String)
        da.SelectCommand.Parameters("ban").Value = CodigoBanco
        dt.Clear()
        da.Fill(dt)

        If dt.Rows.Count = 1 Then
            dr = dt.Rows(0)
        Else
            dr = Nothing
        End If
    End Sub
    Public ReadOnly Property CBU() As String
        Get
            Return dr("bidnum_0").ToString
        End Get
    End Property
    Public ReadOnly Property AAlias() As String
        Get
            Return dr("xalias_0").ToString
        End Get
    End Property

End Class
