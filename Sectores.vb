Imports System.Data.OracleClient
Imports System.ComponentModel

Public Class Sectores
    Inherits BindingList(Of Sector)

    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
        Adaptadores()
    End Sub
    Private Sub Adaptadores()
        Dim Sql As String

        Sql = "SELECT * FROM xsectores WHERE bpcnum_0 = :bpcnum AND fcyitn_0 = :fcyitn"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("bpcnum", OracleType.VarChar)
        da.SelectCommand.Parameters.Add("fcyitn", OracleType.VarChar)
    End Sub
    Public Sub AbrirSectores(ByVal Cliente As String, ByVal Sucursal As String)
        da.SelectCommand.Parameters("bpcnum").Value = Cliente
        da.SelectCommand.Parameters("fcyitn").Value = Sucursal

        Try
            dt.Clear()
            Me.Clear()

            da.Fill(dt)

            For Each dr As DataRow In dt.Rows
                Dim s As New Sector(cn)
                If s.Abrir(CLng(dr(0))) Then Me.Add(s)
            Next

        Catch ex As Exception
        End Try
    End Sub

End Class
