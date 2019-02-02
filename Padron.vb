Imports System.Data.OracleClient

Public Class Padron
    Private Const ALICUOTA_MAXIMA As Double = 6
    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn

        Dim Sql As String

        Sql = "SELECT * FROM xdbpad WHERE crn_0 = :crn"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("crn", OracleType.VarChar)
    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal Cuit As String)
        Me.New(cn)
        Buscar(Cuit)
    End Sub


    Public Function Buscar(ByVal cuit As String) As Boolean

        da.SelectCommand.Parameters("crn").Value = cuit

        Try
            dt.Clear()
            da.Fill(dt)

            Return dt.Rows.Count > 0

        Catch ex As Exception
        End Try


    End Function

    Public ReadOnly Property AlicuotaPercepcion() As Double
        Get
            Dim dr As DataRow
            Dim a As Double = 0

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                a = CDbl(dr("ibcfepper_0"))
            Else
                a = ALICUOTA_MAXIMA
            End If

            Return a
        End Get
    End Property
    Public ReadOnly Property PercepcionDesde() As Date
        Get
            Dim dr As DataRow
            Dim d As Date

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                d = CDate(dr("ibcfeddeb_0"))
            Else
                d = New Date(Today.Year, Today.Month, 1)
            End If

            Return d
        End Get
    End Property
    Public ReadOnly Property PercepcionHasta() As Date
        Get
            Dim dr As DataRow
            Dim d As Date

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                d = CDate(dr("ibcfedfin_0"))
            Else
                d = New Date(Today.Month, Today.Month, 1)
                d = d.AddYears(10).AddDays(-1)
            End If

            Return d
        End Get
    End Property

End Class
