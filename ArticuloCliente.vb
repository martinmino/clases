Imports System.Data.OracleClient

Public Class ArticuloCliente

    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable


    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String

        Me.cn = cn

        Sql = "SELECT * FROM itmbpc WHERE itmref_0 = :itmref AND bpcnum_0 = :bpcnum"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("itmref", OracleType.VarChar)
        da.SelectCommand.Parameters.Add("bpcnum", OracleType.VarChar)

    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal Articulo As String, ByVal Cliente As String)
        Me.New(cn)
        Abrir(Articulo, Cliente)
    End Sub
    Public Function Abrir(ByVal Articulo As String, ByVal Cliente As String) As Boolean
        dt.Clear()

        da.SelectCommand.Parameters("itmref").Value = Articulo
        da.SelectCommand.Parameters("bpcnum").Value = Cliente

        Try
            dt.Clear()
            da.Fill(dt)

        Catch ex As Exception

        End Try

        Return dt.Rows.Count > 0

    End Function

    Public ReadOnly Property ArticuloCliente() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("itmrefbpc_0").ToString()
        End Get
    End Property
    Public ReadOnly Property DescripcionCliente() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("itmdesbpc_0").ToString()
        End Get
    End Property
    Public ReadOnly Property CantidadPorCaja() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CInt(dr("pckcap_0"))
        End Get
    End Property

End Class