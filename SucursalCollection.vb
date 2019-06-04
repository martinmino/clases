Imports System.Data.OracleClient
Imports System.ComponentModel

Public Class SucursalCollection
    Inherits BindingList(Of Sucursal)

    Private cn As OracleConnection
    Public dt As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
    End Sub

    Public Sub CargarSucursales(ByVal Codigo As String)
        Dim Sql As String
        Dim da As OracleDataAdapter

        Sql = "SELECT bpanum_0, bpaadd_0, bpaaddlig_0, bpaadd_0 || ' - ' || bpaaddlig_0 as direccion FROM bpaddress WHERE bpanum_0 = :bpanum ORDER BY bpaadd_0"
        da = New OracleDataAdapter(Sql, Me.cn)
        da.SelectCommand.Parameters.Add("bpanum", OracleType.VarChar).Value = Codigo

        Try
            da.Fill(dt)

            For Each dr As DataRow In dt.Rows
                Dim s As New Sucursal(cn)
                If s.Abrir(dr("bpanum_0").ToString, dr("bpaadd_0").ToString) Then Me.Add(s)
            Next

        Catch ex As Exception

        End Try

    End Sub
    Public Function SiguienteCodigoSucursal() As String
        Dim i As Integer
        Dim n As Integer = 0
        Dim Existe As Boolean = False

        For i = 1 To 999
            Existe = False

            For Each s As Sucursal In Me
                If IsNumeric(s.Sucursal) Then
                    n = CInt(s.Sucursal)
                    If i = n Then Existe = True
                End If

                If Existe Then Exit For
            Next

            If Not Existe Then Exit For
        Next

        Return Strings.Right("00" & i.ToString, 3)

    End Function

End Class
