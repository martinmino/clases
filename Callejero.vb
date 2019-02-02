Imports System.Data.OracleClient

Public Class Callejero

    Private cn As OracleConnection
    Private dt As New DataTable
    Private da As OracleDataAdapter

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String = ""

        Me.cn = cn

        Sql = "SELECT * FROM xcallejero"
        da = New OracleDataAdapter(Sql, cn)
        CargarCallejero()

    End Sub
    Public Sub CargarCallejero()
        dt.Clear()
        da.Fill(dt)
    End Sub
    Public Function ObtenerCalle(ByVal id As String) As String
        Dim dv As New DataView(dt)
        Dim Calle As String = ""

        dv.RowFilter = "calle_0=" & id

        If dv.Count > 0 Then
            Calle = dv.Item(0).Item(1).ToString()
        End If

        Return Calle

    End Function
End Class
