Imports System.Data.OracleClient

Public Class Impuesto
    Private cn As OracleConnection
    Private dt1 As New DataTable 'TABVAC
    Private dt2 As New DataTable 'TABVAT
    Private da1 As OracleDataAdapter
    Private da2 As OracleDataAdapter

    Public Sub New(ByVal cn As OracleConnection, ByVal Tipo As String, ByVal bpc As Cliente)
        Dim Sql As String

        Me.cn = cn

        Sql = "SELECT * FROM tabvac WHERE vacbpr_0 = :vacbpr AND vacitm_0 = :vacitm"
        da1 = New OracleDataAdapter(Sql, cn)
        With da1.SelectCommand.Parameters
            .Add("vacbpr", OracleType.VarChar)
            .Add("vacitm", OracleType.VarChar)
        End With

        Sql = "SELECT * FROM tabvat WHERE vat_0 = :vat"
        da2 = New OracleDataAdapter(Sql, cn)
        da2.SelectCommand.Parameters.Add("vat", OracleType.VarChar)

        Abrir(Tipo, bpc)

    End Sub
    Public Function Abrir(ByVal Tipo As String, ByVal bpc As Cliente) As Boolean
        dt1.Clear()
        dt2.Clear()

        da1.SelectCommand.Parameters("vacbpr").Value = bpc.RegimenImpuesto
        da1.SelectCommand.Parameters("vacitm").Value = Tipo
        da1.Fill(dt1)

        If dt1.Rows.Count = 1 Then
            da2.SelectCommand.Parameters("vat").Value = CodigoAlicuota
            da2.Fill(dt2)
        End If

    End Function
    Public ReadOnly Property Alicuota() As Double
        Get
            Dim dr As DataRow = dt2.Rows(0)
            Return CDbl(dr("vatrat_0"))
        End Get
    End Property
    Public ReadOnly Property CodigoAlicuota() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("vat_0").ToString
        End Get
    End Property
    Public ReadOnly Property Tipo() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("vacitm_0").ToString
        End Get
    End Property

End Class
