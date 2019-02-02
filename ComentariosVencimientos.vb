Imports System.Data.OracleClient

Public Class ComentariosVencimientos
    Private cn As OracleConnection
    Private da As OracleDataAdapter

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String

        Me.cn = cn

        Sql = "SELECT * FROM xvtoscom where fecha_0 = :p1 and bpcnum_0 = :p2 and fcyitn_0 = :p3"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("p1", OracleType.DateTime)
        da.SelectCommand.Parameters.Add("p2", OracleType.VarChar)
        da.SelectCommand.Parameters.Add("p3", OracleType.VarChar)
        da.InsertCommand = New OracleCommandBuilder(da).GetInsertCommand
        da.UpdateCommand = New OracleCommandBuilder(da).GetUpdateCommand
        da.DeleteCommand = New OracleCommandBuilder(da).GetDeleteCommand

    End Sub
    Private Function Buscar(ByVal Fecha As Date, ByVal Cliente As String, ByVal Sucursal As String) As DataTable
        da.SelectCommand.Parameters("p1").Value = Fecha.Date
        da.SelectCommand.Parameters("p2").Value = Cliente
        da.SelectCommand.Parameters("p3").Value = Sucursal

        Dim dt As New DataTable

        da.Fill(dt)

        Return dt

    End Function
    Public Function getComentario(ByVal Fecha As Date, ByVal Cliente As String, ByVal Sucursal As String) As String
        Dim txt As String = ""
        Dim dt As DataTable
        Dim dr As DataRow

        dt = Buscar(Fecha, Cliente, Sucursal)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            txt = dr("obs_0").ToString
        End If

        dt.Dispose()

        Return txt.Trim

    End Function
    Public Function setComentario(ByVal Fecha As Date, ByVal Cliente As String, ByVal Sucursal As String, ByVal txt As String) As Boolean
        Dim dt As DataTable
        Dim dr As DataRow

        dt = Buscar(Fecha, Cliente, Sucursal)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
            dr("fecha_0") = Fecha.Date
            dr("bpcnum_0") = Cliente
            dr("fcyitn_0") = Sucursal
            dr("obs_0") = IIf(txt.Trim = "", " ", txt)
            dt.Rows.Add(dr)
        Else
            dr = dt.Rows(0)
            dr.BeginEdit()
            dr("obs_0") = IIf(txt.Trim = "", " ", txt)
            dr.EndEdit()

        End If

        If txt.Trim = "" Then dr.Delete()

        da.Update(dt)
        dt.Dispose()

    End Function

End Class