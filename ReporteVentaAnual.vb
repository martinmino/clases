Imports System.Data.OracleClient
Imports System.Collections

Public Class ReporteVentaAnual
    Private cn As OracleConnection
    Public Vendedores As New List(Of String)
    Public usr As Usuario
    Private TipoReporte As Integer = 1
    Public Iva As Boolean = False
    Public Costos As Boolean = False
    Public Presupuesto As Boolean = False

    Private daPpal As OracleDataAdapter
    Private dtPpal As New DataTable

    Public AnoConsulta As Integer = 0
    Private Desde As Date
    Private Hasta As Date

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn

        Adaptadores()
    End Sub
    Public Sub Nuevo(ByVal Anio As Integer, ByVal TipoReporte As Integer)
        AnoConsulta = Anio
        Me.TipoReporte = TipoReporte

        If AnoConsulta < 2009 Then AnoConsulta = Today.Year

        If AnoConsulta = Date.Today.Year Then
            Desde = New Date(Today.Year, 1, 1)
            Hasta = New Date(Today.Year, Today.Month, 1)
        Else
            Desde = New Date(AnoConsulta, 1, 1)
            Hasta = Desde.AddYears(1)
        End If

        AbrirTablaPrincipal()

        Select Case TipoReporte
            Case 1
                ObtenerVentas()

            Case 2
                ObtenerVentasExtintor()

            Case 3
                ObtenerVentasSF()

        End Select

        daPpal.Update(dtPpal)

    End Sub
    Private Sub Adaptadores()
        Dim Sql As String

        Sql = "SELECT * FROM xtmpvtas WHERE usr_0 = :usr_0"
        daPpal = New OracleDataAdapter(Sql, cn)
        daPpal.SelectCommand.Parameters.Add("usr_0", OracleType.VarChar)
        daPpal.InsertCommand = New OracleCommandBuilder(daPpal).GetInsertCommand
        daPpal.UpdateCommand = New OracleCommandBuilder(daPpal).GetUpdateCommand
        daPpal.DeleteCommand = New OracleCommandBuilder(daPpal).GetDeleteCommand

    End Sub
    Private Sub AbrirTablaPrincipal()
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim dr1 As DataRow
        Dim Sql As String = ""

        'Recupero la estructura de la tabla con registros anteriores (si existen)
        daPpal.SelectCommand.Parameters("usr_0").Value = usr.Codigo
        dtPpal.Clear()
        daPpal.Fill(dtPpal)

        'Elimino todos los registros anteriores
        For Each dr1 In dtPpal.Rows
            dr1.Delete()
        Next

        Select Case TipoReporte
            Case 1
                'Recupero las familias 4 y 5
                Sql = "SELECT DISTINCT ' ' AS tsicod_1, ' ' AS tsicod_2, tsicod_3, tsicod_4 "
                Sql &= "FROM itmmaster "
                Sql &= "WHERE tsicod_3 <> ' ' AND tsicod_4 <> ' ' "
                Sql &= "ORDER BY tsicod_3, tsicod_4"

                da = New OracleDataAdapter(Sql, cn)
                da.Fill(dt)

                'Agrego las familias 4 y 5 a la tabla
                For Each dr In dt.Rows
                    dr1 = dtPpal.NewRow
                    dr1("usr_0") = usr.Codigo
                    dr1("tsicod1_0") = " "
                    dr1("tsicod2_0") = " "
                    dr1("tsicod3_0") = dr("tsicod_3")
                    dr1("tsicod4_0") = dr("tsicod_4")
                    dr1("qtydia_0") = 0
                    dr1("predia_0") = 0
                    dr1("impdia_0") = 0
                    dr1("mardia_0") = 0
                    dr1("qtyacu_0") = 0
                    dr1("preacu_0") = 0
                    dr1("impacu_0") = 0
                    dr1("maracu_0") = 0
                    dr1("qtypro_0") = 0
                    dr1("imppro_0") = 0
                    dr1("marpro_0") = 0
                    dr1("qtyhis_0") = 0
                    dr1("prehis_0") = 0
                    dr1("imphis_0") = 0
                    dr1("fecha_0") = Hasta.AddDays(-1)
                    dtPpal.Rows.Add(dr1)
                Next

            Case 2
                'Recupero las familias 4 y 5
                Sql = " SELECT EXTRACT(MONTH FROM accdat_0) AS MES, itm.tsicod_4, itm.tsicod_2, itm.tsicod_1, ' ' AS cant  "
                Sql &= "FROM (sinvoice siv INNER JOIN sinvoiced sid ON (siv.num_0 = sid.num_0)) INNER JOIN itmmaster itm ON (sid.itmref_0 = itm.itmref_0) "
                Sql &= "WHERE EXTRACT(YEAR FROM accdat_0) = :accdat_0 AND sivtyp_0 <> 'PRF' AND tsicod_3 = '301' "
                Sql &= "GROUP BY EXTRACT(MONTH FROM accdat_0), itm.tsicod_4, itm.tsicod_2, itm.tsicod_1 "
                Sql &= "ORDER BY mes, tsicod_4, tsicod_2, tsicod_1"

                da = New OracleDataAdapter(Sql, cn)
                da.SelectCommand.Parameters.Add("accdat_0", OracleType.Number).Value = AnoConsulta
                da.Fill(dt)

                'Agrego las familias 4 y 5 a la tabla
                For Each dr In dt.Rows
                    dr1 = dtPpal.NewRow
                    dr1("usr_0") = usr.Codigo
                    dr1("tsicod1_0") = dr("tsicod_1")
                    dr1("tsicod2_0") = dr("tsicod_2")
                    dr1("tsicod3_0") = dr("mes")
                    dr1("tsicod4_0") = dr("tsicod_4")
                    dr1("qtydia_0") = 0
                    dr1("predia_0") = dr("mes")
                    dr1("impdia_0") = 0
                    dr1("mardia_0") = 0
                    dr1("qtyacu_0") = 0
                    dr1("preacu_0") = 0
                    dr1("impacu_0") = 0
                    dr1("maracu_0") = 0
                    dr1("qtypro_0") = 0
                    dr1("imppro_0") = 0
                    dr1("marpro_0") = 0
                    dr1("qtyhis_0") = 0
                    dr1("prehis_0") = 0
                    dr1("imphis_0") = 0
                    dr1("fecha_0") = Hasta.AddDays(-1)
                    dtPpal.Rows.Add(dr1)
                Next

            Case 3
                'Recupero las familias 2, 3, 4 y 5
                Sql = "SELECT DISTINCT tsicod_1, tsicod_2, tsicod_3, tsicod_4 "
                Sql &= "FROM itmmaster "
                Sql &= "WHERE tsicod_1 BETWEEN '133' AND '195' AND "
                Sql &= "      tsicod_2 <> ' ' AND "
                Sql &= "      tsicod_3 = '304' AND "
                Sql &= "      tsicod_4 <> ' ' "
                Sql &= "ORDER BY tsicod_4, tsicod_2, tsicod_1"

                da = New OracleDataAdapter(Sql, cn)
                da.Fill(dt)

                'Agrego las familias 2, 3, 4 y 5 a la tabla
                For Each dr In dt.Rows
                    dr1 = dtPpal.NewRow
                    dr1("usr_0") = usr.Codigo
                    dr1("tsicod1_0") = dr("tsicod_1")
                    dr1("tsicod2_0") = dr("tsicod_2")
                    dr1("tsicod3_0") = dr("tsicod_3")
                    dr1("tsicod4_0") = dr("tsicod_4")
                    dr1("qtydia_0") = 0
                    dr1("predia_0") = 0
                    dr1("impdia_0") = 0
                    dr1("mardia_0") = 0
                    dr1("qtyacu_0") = 0
                    dr1("preacu_0") = 0
                    dr1("impacu_0") = 0
                    dr1("maracu_0") = 0
                    dr1("qtypro_0") = 0
                    dr1("imppro_0") = 0
                    dr1("marpro_0") = 0
                    dr1("qtyhis_0") = 0
                    dr1("prehis_0") = 0
                    dr1("imphis_0") = 0
                    dr1("fecha_0") = Hasta.AddDays(-1)
                    dtPpal.Rows.Add(dr1)
                Next
        End Select

    End Sub
    Private Sub ObtenerVentas()
        Dim Sql As String
        Dim dt As New DataTable
        Dim da As OracleDataAdapter
        Dim dv As DataView
        Dim dr As DataRow
        Dim itm As New Articulo(cn)
        Dim Costo As Double = 0

        'Consulta para recuperar ventas del AÑO actual y del año anterior
        Sql = "SELECT  qty_0 * sns_0 AS cant, qty_0 * sns_0 * netpri_0 * ratcur_0 AS punit, "
        Sql &= "       {amtatilin_0} * sns_0 * ratcur_0 AS impii, "
        Sql &= "       itm.tsicod_3, itm.tsicod_4, accdat_0, rep1_0, sid.itmref_0 "
        Sql &= "FROM (sinvoice siv INNER JOIN sinvoiced sid ON (siv.num_0 = sid.num_0)) INNER JOIN itmmaster itm ON (sid.itmref_0 = itm.itmref_0) "
        Sql &= "WHERE accdat_0 >= :desde AND "
        Sql &= "      accdat_0 < :hasta AND "
        Sql &= "      sivtyp_0 <> 'PRF' "
        Sql &= "ORDER BY accdat_0"

        If Iva Then
            Sql = Strings.Replace(Sql, "{amtatilin_0}", "amtatilin_0")
        Else
            Sql = Strings.Replace(Sql, "{amtatilin_0}", "amtnotlin_0")
        End If

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("desde", OracleType.DateTime)
        da.SelectCommand.Parameters.Add("hasta", OracleType.DateTime)

        da.SelectCommand.Parameters("desde").Value = Desde
        da.SelectCommand.Parameters("hasta").Value = Hasta
        da.Fill(dt)

        If Not Presupuesto Then
            da.SelectCommand.Parameters("desde").Value = Desde.AddYears(-1)
            da.SelectCommand.Parameters("hasta").Value = Hasta.AddYears(-1)
            da.Fill(dt)
        End If

        dv = New DataView(dtPpal)

        'CALCULO DE VENTAS DEL AÑO Y AÑO ANTERIOR
        For Each dr In dt.Rows
            'Proceso el registro si el vendedor está chequeado
            If Vendedores.IndexOf(dr("rep1_0").ToString) > -1 Then

                'Filtro la familia a la que voy a sumar las cantidades
                dv.RowFilter = "tsicod3_0 = '" & dr("tsicod_3").ToString & "' AND tsicod4_0 = '" & dr("tsicod_4").ToString & "'"

                If dv.Count = 1 Then 'Encontre la familia
                    With dv.Item(0).Row
                        .BeginEdit()
                        'Calculo venta de año consultado
                        If CDate(dr("accdat_0")).Year = AnoConsulta Then

                            .Item("qtydia_0") = CDbl(.Item("qtydia_0")) + CDbl(dr("cant"))
                            '.Item("predia_0") = CDbl(.Item("predia_0")) + CDbl(dr("punit"))
                            .Item("impdia_0") = CDbl(.Item("impdia_0")) + CDbl(dr("impii"))

                            .Item("qtyacu_0") = CDbl(.Item("qtyacu_0")) + CDbl(dr("cant"))
                            '.Item("preacu_0") = CDbl(.Item("preacu_0")) + CDbl(dr("punit"))
                            .Item("impacu_0") = CDbl(.Item("impacu_0")) + CDbl(dr("impii"))

                            If Costos Then
                                If itm.Abrir(dr("itmref_0").ToString) Then
                                    Costo = CDbl(dr("cant")) * itm.Costo
                                    .Item("mardia_0") = CDbl(.Item("mardia_0")) + Costo
                                End If
                            End If

                        Else
                            .Item("qtyhis_0") = CDbl(.Item("qtyhis_0")) + CDbl(dr("cant"))
                            '.Item("prehis_0") = CDbl(.Item("prehis_0")) + CDbl(dr("punit"))
                            .Item("imphis_0") = CDbl(.Item("imphis_0")) + CDbl(dr("impii"))

                        End If
                        .EndEdit()
                    End With

                End If

            End If
        Next

        If Presupuesto Then
            AplicarPresupuesto(Hasta)
        End If

        'Elimino registros sin cantidades
        For i As Integer = dtPpal.Rows.Count - 1 To 0 Step -1
            dr = dtPpal.Rows(i)

            If dr.RowState <> DataRowState.Deleted Then

                If CInt(dr("qtydia_0")) = 0 And CInt(dr("qtyacu_0")) = 0 And CInt(dr("qtyhis_0")) = 0 Then
                    'Elimino registro sin cantidades
                    dr.Delete()
                End If

            End If

        Next

    End Sub
    Private Sub ObtenerVentasSF()
        Dim Sql As String
        Dim dt As New DataTable
        Dim da As OracleDataAdapter
        Dim dv As DataView
        Dim dr As DataRow
        Dim itm As New Articulo(cn)
        Dim Costo As Double = 0

        'Consulta para recuperar ventas del AÑO actual y del año anterior
        Sql = "SELECT  qty_0 * sns_0 AS cant, "
        Sql &= "       qty_0 * sns_0 * netpri_0 * ratcur_0 AS punit, "
        Sql &= "       {amtatilin_0} * sns_0 * ratcur_0 AS impii, "
        Sql &= "       itm.tsicod_1, "
        Sql &= "       itm.tsicod_2, "
        Sql &= "       itm.tsicod_3, "
        Sql &= "       itm.tsicod_4, "
        Sql &= "       accdat_0, "
        Sql &= "       rep1_0, "
        Sql &= "       sid.itmref_0 "
        Sql &= "FROM sinvoice siv INNER JOIN "
        Sql &= "     sinvoiced sid ON (siv.num_0 = sid.num_0) INNER JOIN "
        Sql &= "     itmmaster itm ON (sid.itmref_0 = itm.itmref_0) "
        Sql &= "WHERE accdat_0 >= :desde AND "
        Sql &= "      accdat_0 < :hasta AND "
        Sql &= "      sivtyp_0 <> 'PRF' "
        Sql &= "ORDER BY accdat_0"

        If Iva Then
            Sql = Strings.Replace(Sql, "{amtatilin_0}", "amtatilin_0")
        Else
            Sql = Strings.Replace(Sql, "{amtatilin_0}", "amtnotlin_0")
        End If

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("desde", OracleType.DateTime)
        da.SelectCommand.Parameters.Add("hasta", OracleType.DateTime)

        da.SelectCommand.Parameters("desde").Value = Desde
        da.SelectCommand.Parameters("hasta").Value = Hasta
        da.Fill(dt)

        If Not Presupuesto Then
            da.SelectCommand.Parameters("desde").Value = Desde.AddYears(-1)
            da.SelectCommand.Parameters("hasta").Value = Hasta.AddYears(-1)
            da.Fill(dt)
        End If

        dv = New DataView(dtPpal)

        'CALCULO DE VENTAS DEL AÑO Y AÑO ANTERIOR
        For Each dr In dt.Rows
            'Proceso el registro si el vendedor está chequeado
            If Vendedores.IndexOf(dr("rep1_0").ToString) > -1 Then

                'Filtro la familia a la que voy a sumar las cantidades
                dv.RowFilter = "tsicod1_0 = '" & dr("tsicod_1").ToString & "' AND " _
                             & "tsicod2_0 = '" & dr("tsicod_2").ToString & "' AND " _
                             & "tsicod3_0 = '" & dr("tsicod_3").ToString & "' AND " _
                             & "tsicod4_0 = '" & dr("tsicod_4").ToString & "'"

                If dv.Count = 1 Then 'Encontre la familia
                    With dv.Item(0).Row
                        .BeginEdit()
                        'Calculo venta de año consultado
                        If CDate(dr("accdat_0")).Year = AnoConsulta Then

                            .Item("qtydia_0") = CDbl(.Item("qtydia_0")) + CDbl(dr("cant"))
                            '.Item("predia_0") = CDbl(.Item("predia_0")) + CDbl(dr("punit"))
                            .Item("impdia_0") = CDbl(.Item("impdia_0")) + CDbl(dr("impii"))

                            .Item("qtyacu_0") = CDbl(.Item("qtyacu_0")) + CDbl(dr("cant"))
                            '.Item("preacu_0") = CDbl(.Item("preacu_0")) + CDbl(dr("punit"))
                            .Item("impacu_0") = CDbl(.Item("impacu_0")) + CDbl(dr("impii"))

                            If Costos Then
                                If itm.Abrir(dr("itmref_0").ToString) Then
                                    Costo = CDbl(dr("cant")) * itm.Costo
                                    .Item("mardia_0") = CDbl(.Item("mardia_0")) + Costo
                                End If
                            End If

                        Else
                            .Item("qtyhis_0") = CDbl(.Item("qtyhis_0")) + CDbl(dr("cant"))
                            '.Item("prehis_0") = CDbl(.Item("prehis_0")) + CDbl(dr("punit"))
                            .Item("imphis_0") = CDbl(.Item("imphis_0")) + CDbl(dr("impii"))

                        End If
                        .EndEdit()
                    End With

                End If

            End If
        Next

        If Presupuesto Then
            AplicarPresupuesto(Hasta)
        End If

        'Elimino registros sin cantidades
        For i As Integer = dtPpal.Rows.Count - 1 To 0 Step -1
            dr = dtPpal.Rows(i)

            If dr.RowState <> DataRowState.Deleted Then

                If CInt(dr("qtydia_0")) = 0 And CInt(dr("qtyacu_0")) = 0 And CInt(dr("qtyhis_0")) = 0 Then
                    'Elimino registro sin cantidades
                    dr.Delete()
                End If

            End If

        Next

    End Sub
    Private Sub ObtenerVentasExtintor()
        'Obtiene las cantidades vendidas por familias por mes
        Dim Sql As String
        Dim dt As New DataTable
        Dim da As OracleDataAdapter
        Dim dv As DataView
        Dim dr1 As DataRow

        Sql = "SELECT EXTRACT(MONTH FROM accdat_0) AS MES, itm.tsicod_4, itm.tsicod_2, itm.tsicod_1, sum(qty_0 * sns_0) AS cant, REP1_0  "
        Sql &= "FROM (sinvoice siv INNER JOIN sinvoiced sid ON (siv.num_0 = sid.num_0)) INNER JOIN itmmaster itm ON (sid.itmref_0 = itm.itmref_0) "
        Sql &= "WHERE EXTRACT(YEAR FROM accdat_0) = :accdat_0 AND sivtyp_0 <> 'PRF' AND tsicod_3 = '301' "
        Sql &= "GROUP BY EXTRACT(MONTH FROM accdat_0), itm.tsicod_4, itm.tsicod_2, itm.tsicod_1 , REP1_0 "
        Sql &= "ORDER BY mes, tsicod_4, tsicod_2, tsicod_1"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("accdat_0", OracleType.Number).Value = AnoConsulta
        da.Fill(dt)
        da.Dispose()

        dv = New DataView(dtPpal)
        For Each dr1 In dt.Rows
            'Proceso el registro si el vendedor está chequeado
            If Vendedores.IndexOf(dr1("rep1_0").ToString) > -1 Then

                'Filtro la familia a la que voy a sumar las cantidades
                dv.RowFilter = "tsicod1_0 = '" & dr1("tsicod_1").ToString & "' AND tsicod2_0 = '" & dr1("tsicod_2").ToString & "' AND tsicod3_0 = '" & dr1("mes").ToString & "' AND tsicod4_0 = '" & dr1("tsicod_4").ToString & "'"

                If dv.Count = 1 Then 'Encontre la familia
                    With dv.Item(0).Row
                        .BeginEdit()
                        'Calculo venta de año consultado

                        .Item("qtydia_0") = CDbl(.Item("qtydia_0")) + CDbl(dr1("cant"))
                        .EndEdit()
                    End With

                End If

            End If

        Next

    End Sub
    Private Sub AplicarPresupuesto(ByVal Fecha As Date)
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim dv As DataView

        Fecha = Fecha.AddDays(-1)

        Sql = "select tsicod3_0, tsicod4_0, sum(qty_0) as qty, sum(imp_0) as imp "
        Sql &= "from presupues "
        Sql &= "where ano_0 = :ano and "
        Sql &= "	  mes_0 <= :mes "
        Sql &= "group by tsicod3_0, tsicod4_0"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("ano", OracleType.Number).Value = Fecha.Year
        da.SelectCommand.Parameters.Add("mes", OracleType.Number).Value = Fecha.Month

        da.Fill(dt)
        dv = New DataView(dt)

        For Each dr As DataRow In dtPpal.Rows
            If dr.RowState = DataRowState.Deleted Then Continue For

            Dim f As String = "tsicod3_0 = '{3}' and tsicod4_0 = '{4}'"
            f = f.Replace("{3}", dr("tsicod3_0").ToString)
            f = f.Replace("{4}", dr("tsicod4_0").ToString)

            dv.RowFilter = f

            If dv.Count = 0 Then
                dr.BeginEdit()
                dr("qtyhis_0") = 0
                dr("prehis_0") = 0
                dr("imphis_0") = 0
                dr.EndEdit()
            Else
                Dim q As Double = CDbl(dv.Item(0).Item(2))
                Dim i As Double = CDbl(dv.Item(0).Item(3))

                dr.BeginEdit()
                dr("qtyhis_0") = q
                dr("imphis_0") = i
                dr.EndEdit()
            End If

        Next

    End Sub

End Class