Imports System.Data.OracleClient

Public Class Devolucion
    Private cn As OracleConnection
    Private dah As OracleDataAdapter
    Private dad As OracleDataAdapter
    Private dth As New DataTable
    Private dtd As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn

        Dim Sql As String = ""

        Sql = "SELECT * FROM sreturn WHERE srhnum_0 = :srhnum"
        dah = New OracleDataAdapter(Sql, cn)
        dah.SelectCommand.Parameters.Add("srhnum", OracleType.VarChar)

        Sql = "SELECT * FROM sreturnd WHERE srhnum_0 = :srhnum"
        dad = New OracleDataAdapter(Sql, cn)
        dad.SelectCommand.Parameters.Add("srhnum", OracleType.VarChar)

    End Sub
    Public Function Abrir(ByVal NumeroDevolucion As String) As Boolean
        dah.SelectCommand.Parameters("srhnum").Value = NumeroDevolucion
        dad.SelectCommand.Parameters("srhnum").Value = NumeroDevolucion

        Try
            dth.Clear()
            dtd.Clear()

            dah.Fill(dth)
            dad.Fill(dtd)

        Catch ex As Exception

        End Try

        Return dth.Rows.Count > 0

    End Function

    Public ReadOnly Property Numero() As String
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return dr("srhnum_0").ToString
        End Get
    End Property
    Public ReadOnly Property Cliente() As Cliente
        Get
            Dim dr As DataRow = dth.Rows(0)
            Dim bpc As New Cliente(cn)
            bpc.Abrir(dr("bpcord_0").ToString)

            Return bpc
        End Get
    End Property
    Public ReadOnly Property Sucursal() As Sucursal
        Get
            Dim dr As DataRow = dth.Rows(0)
            Dim bpa As New Sucursal(cn)

            bpa.Abrir(dr("bpcord_0").ToString, dr("bpaadd_0").ToString)
            Return bpa

        End Get
    End Property
    Public ReadOnly Property Detalle() As DataTable
        Get
            Return dtd
        End Get
    End Property

End Class
