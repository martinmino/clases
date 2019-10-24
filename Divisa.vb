Imports System.Data.OracleClient

Public Class Divisa
    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable
    Private dr As DataRow

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String = ""
        Me.cn = cn

        Sql = "SELECT * from tabcur WHERE cur_0 = :cur"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("cur", OracleType.VarChar)
    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal Codigo As String)
        Me.New(cn)
        Abrir(Codigo)
    End Sub
    Public Sub Abrir(ByVal Codigo As String)
        da.SelectCommand.Parameters("cur").Value = Codigo
        dt.Clear()
        da.Fill(dt)

        If dt.Rows.Count = 1 Then
            dr = dt.Rows(0)
        Else
            dr = Nothing
        End If
    End Sub
    Public ReadOnly Property Codigo() As String
        Get
            Return dr("cur_0").ToString
        End Get
    End Property
    Public ReadOnly Property Nombre() As String
        Get
            Return dr("curdes_0").ToString
        End Get
    End Property
    Public ReadOnly Property CodigoAFIP() As String
        Get
            Return dr("isocod_0").ToString
        End Get
    End Property
End Class
