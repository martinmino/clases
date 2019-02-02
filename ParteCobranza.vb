Imports System.Data.OracleClient
Imports CrystalDecisions.CrystalReports.Engine

Public Class ParteCobranza
    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable
    Private Factura As String = ""

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
        Adaptadores()
    End Sub
    Private Sub Adaptadores()
        Dim Sql As String = ""

        Sql = "SELECT * FROM gaccdudate WHERE num_0 = :num"
        da = New OracleDataAdapter(Sql, cn)

        da.SelectCommand.Parameters.Add("num", OracleType.VarChar)
        da.UpdateCommand = New OracleCommandBuilder(da).GetUpdateCommand

    End Sub
    Public Sub Abrir(ByVal NumeroFactura As String)
        Me.Factura = NumeroFactura

        da.SelectCommand.Parameters("num").Value = NumeroFactura

        dt.Clear()
        da.Fill(dt)

    End Sub
    Public Sub Marcar(ByVal Fecha As Date, ByVal Cobrador As String)
        For Each dr As DataRow In dt.Rows
            dr.BeginEdit()
            dr("ybprpth_0") = Cobrador
            dr("ycobdat_0") = Fecha
            dr.EndEdit()
        Next
    End Sub
    Public Sub Grabar()
        Dim i As Integer = da.Update(dt)
    End Sub
    Public Sub Imprimir(ByVal Reporte As String)
        Dim rpt As New ReportDocument

        rpt.Load(Reporte)
        rpt.SetParameterValue("FACTURA", Factura)

        rpt.PrintToPrinter(1, False, 1, 100)

    End Sub

    Public ReadOnly Property Cobrador() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                txt = dr("ybprpth_0").ToString
            End If

            Return txt

        End Get
    End Property

End Class
