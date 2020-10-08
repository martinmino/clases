Imports System.Data.OracleClient
Imports System.ComponentModel

Public Class SucursalCollection
    Inherits BindingList(Of Sucursal)

    Private cn As OracleConnection
    Public dt1 As New DataTable 'Sucursal
    Public dt2 As New DataTable 'Sucursal Entregas

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
    End Sub

    Public Sub CargarSucursales(ByVal Codigo As String)
        Dim Sql As String
        Dim da1 As OracleDataAdapter
        Dim da2 As OracleDataAdapter

        'Sucursales
        Sql = "SELECT * FROM bpaddress WHERE bpanum_0 = :bpanum ORDER BY bpaadd_0"
        da1 = New OracleDataAdapter(Sql, Me.cn)
        da1.SelectCommand.Parameters.Add("bpanum", OracleType.VarChar).Value = Codigo

        Sql = "SELECT * FROM bpdlvcust WHERE bpcnum_0 = :bpcnum_0"
        da2 = New OracleDataAdapter(Sql, cn)
        da2.SelectCommand.Parameters.Add("bpcnum_0", OracleType.VarChar).Value = Codigo

        Try
            da1.Fill(dt1)
            da2.Fill(dt2)

            Dim dr1 As DataRow 'Sucursal
            Dim dr2 As DataRow 'SucursalEntrega

            For Each dr1 In dt1.Rows
                'Busco el registro de sucursal-entrega
                dr2 = Nothing
                For Each drx As DataRow In dt2.Rows
                    If drx("bpaadd_0").ToString = dr1("bpaadd_0").ToString Then
                        dr2 = drx
                        Exit For
                    End If
                Next

                Dim s As New Sucursal(cn)
                If s.Abrir(dr1, dr2) Then Me.Add(s)
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
