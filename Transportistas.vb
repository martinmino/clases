Imports System.Data.OracleClient

Public Class Transportistas

    Private da As OracleDataAdapter
    Private dt As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Dim sql As String
        sql = "SELECT * "
        sql &= "FROM bpcarrier "
        sql &= "WHERE bptnum_0 = :bptnum "
        da = New OracleDataAdapter(sql, cn)
        da.SelectCommand.Parameters.Add("bptnum", OracleType.VarChar)
    End Sub
    Public Function Abrir(ByVal Empresa As String, ByVal Patente As String) As Boolean
        da.SelectCommand.Parameters("bptnum").Value = Empresa
        Try
            dt.Clear()
            da.Fill(dt)
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function
    Public ReadOnly Property Codigo() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("bptnum_0").ToString
        End Get
    End Property
    Public ReadOnly Property tarifa() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CInt(dr("fxdamt_0"))
        End Get
    End Property
    Public ReadOnly Property Nombre() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("bptnam_0").ToString
        End Get
    End Property
End Class

