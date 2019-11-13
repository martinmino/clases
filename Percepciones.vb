Imports System.Data.OracleClient

Public Class Percepciones

    Private cn As OracleConnection
    Private dt As DataTable
    Private da As OracleDataAdapter
    Private bpc As Cliente

    Public Sub New(ByVal cn As OracleConnection, ByVal bpc As Cliente, Optional ByVal Tipo As Integer = 1)
        Me.cn = cn
        Me.bpc = bpc

        Dim Sql As String
        Sql = "SELECT * "
        Sql &= "FROM xbprinfo "
        Sql &= "WHERE bprnum_0 = :bprnum and "
        Sql &= "      bprtyp_0 = :bprtyp"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("bprnum", OracleType.VarChar)
        da.SelectCommand.Parameters.Add("bprtyp", OracleType.Number)
        da.InsertCommand = New OracleCommandBuilder(da).GetInsertCommand
        da.UpdateCommand = New OracleCommandBuilder(da).GetUpdateCommand
        da.DeleteCommand = New OracleCommandBuilder(da).GetDeleteCommand

        'Cargo Percepciones del tercero
        dt = New DataTable

        Try
            da.SelectCommand.Parameters("bprnum").Value = bpc.Codigo
            da.SelectCommand.Parameters("bprtyp").Value = bpc.Tipo
            da.Fill(dt)

        Catch ex As Exception

        End Try

    End Sub

    Public Sub Agregar(ByVal Percepcion As String, ByVal Alicuota As Double, ByVal Desde As Date, ByVal Hasta As Date)
        Dim dr As DataRow
        Dim i As Integer

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow

            dr("bprnum_0") = bpc.Codigo
            dr("bprtyp_0") = 1
            dr("satsta_0") = 1
            For i = 0 To 29
                dr("sattax_" & i.ToString) = " "
                dr("satgrp_" & i.ToString) = " "
                dr("rtzgrp_" & i.ToString) = " "
                dr("satrat_" & i.ToString) = 0
                dr("satvdatdeb_" & i.ToString) = #12/31/1599#
                dr("satvdatfin_" & i.ToString) = #12/31/1599#
                dr("exctax_" & i.ToString) = " "
                dr("excvattyp_" & i.ToString) = 0
                dr("excert_" & i.ToString) = " "
                dr("excrat_" & i.ToString) = 0
                dr("excdat_" & i.ToString) = #12/31/1599#
                dr("xregib_" & i.ToString) = 0
            Next
            dr("mtbrtzflg_0") = 0
            dr("credat_0") = Date.Today
            dr("creusr_0") = " "
            dr("upddat_0") = #12/31/1599#
            dr("updusr_0") = " "
            dr("xfceflg_0") = 0
            dr("xfcemonto_0") = 0

            dt.Rows.Add(dr)

        End If

        dr = dt.Rows(0)

        For i = 0 To 29

            If dr("sattax_" & i.ToString).ToString.Trim = "" Then
                dr("sattax_" & i.ToString) = Percepcion
                dr("satrat_" & i.ToString) = Alicuota
                dr("satvdatdeb_" & i.ToString) = Desde
                dr("satvdatfin_" & i.ToString) = Hasta
                dr("xregib_" & i.ToString) = 4

                bpc.AgregarPercepcion(i, Percepcion, Alicuota)

                Exit For
            End If
        Next

    End Sub
    Public Sub Grabar()
        Try
            da.Update(dt)
        Catch ex As Exception

        End Try
    End Sub
    Public ReadOnly Property Alicuota(ByVal Codigo As String) As Double
        Get
            Dim dr As DataRow
            Dim a As Double = 0

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                For i = 0 To 29
                    If dr("sattax_" & i.ToString).ToString.Trim = Codigo Then
                        a = CDbl(dr("satrat_" & i.ToString))
                    End If
                Next
            End If

            Return a
        End Get

    End Property
    Public ReadOnly Property Existe(ByVal Codigo As String) As Boolean
        Get
            Dim dr As DataRow
            Dim flg As Boolean = False

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                For i = 0 To 29
                    If dr("sattax_" & i.ToString).ToString.Trim = Codigo Then
                        flg = True
                        Exit For
                    End If
                Next
            End If

            Return flg

        End Get
    End Property

End Class