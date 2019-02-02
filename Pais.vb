Imports System.Data.OracleClient

Public Class Pais
    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String

        Me.cn = cn

        Sql = "SELECT * FROM tabcountry WHERE cry_0 = :cry_0"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("cry_0", OracleType.VarChar)

    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal Codigo As String)
        Me.New(cn)
        Abrir(Codigo)
    End Sub
    Public Function Abrir(ByVal Codigo As String) As Boolean
        da.SelectCommand.Parameters("cry_0").Value = Codigo
        dt.Clear()

        Try
            da.Fill(dt)

        Catch ex As Exception

        End Try

        Return dt.Rows.Count > 0

    End Function

    Public ReadOnly Property Codigo() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("cry_0").ToString
        End Get
    End Property
    Public ReadOnly Property Nombre() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("crynam_0").ToString
        End Get
    End Property

End Class
