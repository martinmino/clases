Imports System.Data.OracleClient

Public Class Planta
    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Dim dt As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
        Dim Sql As String

        Sql = "SELECT * FROM facility WHERE fcy_0 = :fcy"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("fcy", OracleType.VarChar)

    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal fcy As String)
        Me.New(cn)
        Abrir(fcy)
    End Sub
    Public Function Abrir(ByVal fcy As String) As Boolean
        da.SelectCommand.Parameters("fcy").Value = fcy
        dt.Clear()
        da.Fill(dt)
        Return dt.Rows.Count > 0
    End Function

    Public ReadOnly Property CodigoPlanta() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("fcy_0").ToString
        End Get
    End Property
    Public ReadOnly Property NombrePlanta() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("fcynam_0").ToString
        End Get
    End Property
    Public ReadOnly Property SociedadPlanta() As Sociedad
        Get
            Dim cpy As New Sociedad(cn)
            Dim dr As DataRow = dt.Rows(0)

            cpy.abrir(dr("legcpy_0").ToString)

            Return cpy
        End Get
    End Property

End Class
