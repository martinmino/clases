Imports System.Data.OracleClient

Public Class Empleados

    'Implements IDisposable

    Private da As OracleDataAdapter
    Private dt As New DataTable
    Private dr As DataRow
    Private disposedValue As Boolean = False        ' Para detectar llamadas redundantes
    Private cn As OracleConnection

    'NEW
    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn

    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal Codigo As String)
        Me.New(cn)

        Abrir(Codigo)
    End Sub

    Public Overridable Function Abrir(ByVal Codigo As String) As Boolean
        Dim Sql As String = "SELECT * FROM xempleados WHERE legajo_0 = :legajo_0"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("legajo_0", OracleType.VarChar).Value = Codigo

        dt.Clear()
        da.Fill(dt)

        If dt.Rows.Count = 1 Then dr = dt.Rows(0)

        Return dt.Rows.Count <> 0

    End Function
    Public ReadOnly Property nombre() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("empnam_0").ToString
        End Get
    End Property
    Public ReadOnly Property Apellido() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("empape_0").ToString
        End Get
    End Property
End Class
