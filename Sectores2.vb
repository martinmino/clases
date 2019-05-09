Imports System.Data.OracleClient
Imports System.ComponentModel

Public Class Sectores2
    Inherits BindingList(Of Sector2)

    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
        Adaptadores()
    End Sub
    Private Sub Adaptadores()
        Dim Sql As String

        Sql = "SELECT * FROM xsectores2 WHERE bpcnum_0 = :bpcnum AND fcyitn_0 = :fcyitn"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("bpcnum", OracleType.VarChar)
        da.SelectCommand.Parameters.Add("fcyitn", OracleType.VarChar)
    End Sub
    Public Sub Abrir(ByVal Cliente As String, ByVal Sucursal As String)
        da.SelectCommand.Parameters("bpcnum").Value = Cliente
        da.SelectCommand.Parameters("fcyitn").Value = Sucursal

        Try
            dt.Clear()
            Me.Clear()

            da.Fill(dt)

            ArmarColeccion(dt)

        Catch ex As Exception
        End Try
    End Sub
    Public Sub Grabar()
        For Each s As Sector2 In Me
            s.Grabar()
        Next
    End Sub
    Private Sub ArmarColeccion(ByVal dt As DataTable)
        Me.Clear()

        For Each dr As DataRow In dt.Rows
            Dim s As New Sector2(cn)
            If s.Abrir(dr) Then Me.Add(s)
        Next
    End Sub

End Class
