Imports System.Data.OracleClient

Public Class Camioneta
    Private da As OracleDataAdapter
    Private dt As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Dim sql As String

        sql = "SELECT * "
        sql &= "FROM zunitrans "
        sql &= "WHERE bptnum_0 = :bptnum AND patnum_0 = :patnum"

        da = New OracleDataAdapter(sql, cn)
        With da.SelectCommand.Parameters
            .Add("bptnum", OracleType.VarChar)
            .Add("patnum", OracleType.VarChar)
        End With

    End Sub
    Public Function Abrir(ByVal Empresa As String, ByVal Patente As String) As Boolean
        da.SelectCommand.Parameters("bptnum").Value = Empresa
        da.SelectCommand.Parameters("patnum").Value = Patente

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
    Public ReadOnly Property Patente() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("patnum_0").ToString
        End Get
    End Property
    Public ReadOnly Property Sector() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CInt(dr("xsector_0"))
        End Get
    End Property
    Public ReadOnly Property Acomp() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CInt(dr("acomp_0"))
        End Get
    End Property

End Class