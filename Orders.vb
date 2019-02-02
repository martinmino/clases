Imports System.Data.OracleClient

Public Class Orders
    Private cn As OracleConnection
    Private WithEvents da As OracleDataAdapter
    Private dt As New DataTable
    Private cpy As Sociedad

    Public Sub New(ByVal cn As OracleConnection, ByVal cpy As Sociedad)
        Me.cn = cn

        Dim Sql As String

        Sql = "SELECT * FROM orders"
        da = New OracleDataAdapter(Sql, cn)

        da.InsertCommand = New OracleCommandBuilder(da).GetInsertCommand
        da.DeleteCommand = New OracleCommandBuilder(da).GetDeleteCommand
        da.UpdateCommand = New OracleCommandBuilder(da).GetUpdateCommand

        da.FillSchema(dt, SchemaType.Mapped)

        Nuevo(cpy)

    End Sub
    Public Sub Nuevo(ByVal cpy As Sociedad)
        dt.Clear()
        Me.cpy = cpy
    End Sub
    Public Sub AgregarNecesidad(ByVal dt1 As DataTable)
        Dim dr As DataRow

        For Each dr In dt1.Rows
            AgregarNecesidad(dr)
        Next

    End Sub
    Private Sub AgregarNecesidad(ByVal dr1 As DataRow)
        Dim dr As DataRow = dt.NewRow
        Dim id As String = NuevoID(cpy)

        dr("wiptyp_0") = 1
        dr("wipsta_0") = 1
        dr("wipnum_0") = id
        dr("itmref_0") = dr1("itmref_0").ToString
        dr("orifcy_0") = dr1("salfcy_0").ToString
        dr("stofcy_0") = dr1("stofcy_0").ToString
        dr("pjt_0") = " "
        dr("bprnum_0") = dr1("bpcord_0").ToString
        dr("vcrtyp_0") = 2
        dr("vcrnum_0") = dr1("sohnum_0").ToString
        dr("vcrlin_0") = CLng(dr1("soplin_0"))
        dr("vcrseq_0") = CLng(dr1("soqseq_0"))
        dr("strdat_0") = Date.Today
        dr("enddat_0") = Date.Today
        dr("extqty_0") = CLng(dr1("qty_0"))
        dr("cplqty_0") = 0
        dr("rmnextqty_0") = CLng(dr1("qty_0"))
        dr("allqty_0") = 0
        dr("shtqty_0") = 0
        dr("wortiaqty_0") = 0
        dr("fmi_0") = 1
        dr("mrpmes_0") = 1
        dr("mrpdat_0") = #12/31/1599#
        dr("mrpqty_0") = 0
        dr("vcrtypori_0") = 0
        dr("vcrnumori_0") = " "
        dr("vcrlinori_0") = 0
        dr("vcrseqori_0") = 0
        dr("itmrefori_0") = " "
        dr("bomalt_0") = 0
        dr("bomope_0") = 0
        dr("bomofs_0") = 0
        dr("pio_0") = CInt(dr1("dlvpio_0"))
        dr("ori_0") = 2
        dr("abbfil_0") = "SOQ"
        dr("optflg_0") = 1
        dr("expnum_0") = 1
        dr("credat_0") = Date.Today
        dr("creusr_0") = USER
        dr("upddat_0") = #12/31/1599#
        dr("updusr_0") = " "
        dt.Rows.Add(dr)

        dr1.BeginEdit()
        dr1("demnum_0") = id
        dr1.EndEdit()

    End Sub
    Public Function NuevoID(ByVal cpy As Sociedad) As String
        Dim wip As Numerador = New Numerador(cn, "WIP", cpy.Codigo, CInt(Date.Today.ToString("yy")))
        Dim v2 As String = "WIP-" & cpy.PlantaStock & wip.Periodo.ToString & Strings.Right("0000" & wip.Valor.ToString, 5)

        Return v2

    End Function
    Public Sub Grabar()
        Try
            da.Update(dt)

        Catch ex As Exception

        End Try
    End Sub

    Private Sub da_RowUpdated(ByVal sender As Object, ByVal e As System.Data.OracleClient.OracleRowUpdatedEventArgs) Handles da.RowUpdated
        If e.Status = UpdateStatus.ErrorsOccurred Then Exit Sub
        If e.StatementType <> StatementType.Insert Then Exit Sub

        Dim dr As DataRow = e.Row

        Dim itm As String = dr("itmref_0").ToString
        Dim fcy As String = dr("stofcy_0").ToString
        Dim qty As Long = CLng(dr("extqty_0"))

        itmmvt.Sumar(cn, itm, fcy, qty)

    End Sub

End Class