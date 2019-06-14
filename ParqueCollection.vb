Imports System.Data.OracleClient
Imports System.ComponentModel

Public Class ParqueCollection
    Inherits BindingList(Of Parque)

    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn

        Dim Sql As String

        'Recuperacion de todo el parque del cliente en Sigex
        Sql = "SELECT macnum_0 FROM machines WHERE bpcnum_0 = :bpcnum AND fcyitn_0 = :fcyitn"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("bpcnum", OracleType.VarChar)
        da.SelectCommand.Parameters.Add("fcyitn", OracleType.VarChar)

    End Sub
    Public Sub AbrirParqueCliente(ByVal Cliente As String, ByVal Sucursal As String)
        da.SelectCommand.Parameters("bpcnum").Value = Cliente
        da.SelectCommand.Parameters("fcyitn").Value = Sucursal

        dt.Clear()
        Me.Clear()

        da.Fill(dt)

        For Each dr As DataRow In dt.Rows
            Dim p As New Parque(cn)
            If p.Abrir(dr("macnum_0").ToString) Then Me.Add(p)
        Next

    End Sub

End Class