Imports System.Data.OracleClient
Imports System.ComponentModel

Public Class ParqueCollection
    Inherits BindingList(Of Parque)

    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable

    Private Cliente As String
    Private Sucursal As String

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
        Me.Cliente = Cliente
        Me.Sucursal = Sucursal

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
    Public Function VtoGeneral() As Date
        Return Vencimiento(10)
    End Function
    Public Function VtoPhGeneral() As Date
        Return Vencimiento(20)
    End Function
    Private Function Vencimiento(ByVal Tipo As Integer) As Date
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim dr As DataRow

        Sql = "select ymc.datnext_0, count(datnext_0) as cant "
        Sql &= "from machines mac inner join "
        Sql &= "	 ymacitm ymc on (mac.macnum_0 = ymc.macnum_0) inner join "
        Sql &= "	 bomd bmb on (mac.macpdtcod_0 = bmb.itmref_0 and ymc.cpnitmref_0 = bmb.cpnitmref_0 and bomalt_0 = 99 and bomseq_0 = :tipo) "
        Sql &= "where mac.bpcnum_0 = :cli and "
        Sql &= "	  mac.fcyitn_0 = :suc "
        Sql &= "group by datnext_0 "
        Sql &= "order by cant desc"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("cli", OracleType.VarChar).Value = Cliente
        da.SelectCommand.Parameters.Add("suc", OracleType.VarChar).Value = Sucursal
        da.SelectCommand.Parameters.Add("tipo", OracleType.Number).Value = Tipo

        Try
            da.Fill(dt)
            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)

                If IsDBNull(dr(0)) Then
                    Return Today.Date

                Else
                    Return CDate(dr(0))
                End If

            End If

        Catch ex As Exception
            Return #12/31/1599#

        End Try

    End Function

End Class