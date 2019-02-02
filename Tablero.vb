Imports CrystalDecisions.CrystalReports.Engine
Imports System.Text.RegularExpressions
Imports System.Data.OracleClient
Imports System.Net.Mail
Imports System.IO
Imports System.Data.OleDb

Public Class Tablero
    Private Const COLUMNAS As Integer = 72
    Private Const RPTX3 As String = "\\adonix\Folders\GEOPROD\REPORT\SPA\"
    Private Const DB_USR As String = "GEOPROD"
    Private Const DB_PWD As String = "tiger"
    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable
    Private Fecha As Date
    Private Smtp As New SmtpClient

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String

        Me.cn = cn

        Sql = "SELECT * FROM xtablero WHERE dat_0 = :dat_0"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("dat_0", OracleType.DateTime)
        da.InsertCommand = New OracleCommandBuilder(da).GetInsertCommand
        da.UpdateCommand = New OracleCommandBuilder(da).GetUpdateCommand
        da.DeleteCommand = New OracleCommandBuilder(da).GetDeleteCommand

        With Smtp
            .Host = "mail.matafuegosgeorgia.com"
            .EnableSsl = False
            .UseDefaultCredentials = False
            .Credentials = New Net.NetworkCredential("compumap", "georgia")
        End With

    End Sub
    Public Sub Abrir(ByVal Fecha As Date)
        Dim dr As DataRow

        Me.Fecha = Fecha
        da.SelectCommand.Parameters("dat_0").Value = Fecha

        dt.Clear()
        da.Fill(dt)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
            For i = 0 To COLUMNAS - 1
                If i = 0 Then
                    dr(i) = Fecha
                Else
                    dr(i) = 0
                End If
            Next
            dt.Rows.Add(dr)
        End If

    End Sub
    Public Sub Calcular()
        Dim Feriado As New Feriados(cn)
        Dim dr As DataRow
        Dim f1 As Date 'Fecha inicio
        Dim f2 As Date 'Fecha fin

        f1 = New Date(Fecha.Year, Fecha.Month, 1)
        f1 = f1.AddYears(-1)
        f2 = f1.AddMonths(1).AddDays(-1)

        dr = dt.Rows(0)
        dr.BeginEdit()

        'Calculo de Dias habiles
        dr("habiles_0") = Feriado.DiasHabilesMes(Fecha)
        dr("habiles_1") = Feriado.DiasHabilesHasta(Fecha)


        GoTo saltar

        '***********************************************************************************
        ' SECCION ADMINISTRACION Y FINANZAS
        '***********************************************************************************
        'Cobranzas
        dr("dia_0") = Cobranzas(Fecha, Fecha)
        dr("acu_0") = Cobranzas(New Date(Fecha.Year, Fecha.Month, 1), Fecha)
        dr("his_0") = Cobranzas(f1, f2)
        'Dias en la calle
        dr("dia_1") = DiasCalle()
        dr("acu_1") = 0
        dr("his_1") = 0
        'Dias de pago
        dr("dia_2") = DiasCobro(Fecha)
        dr("acu_2") = 0
        dr("his_2") = 0

        dr("dia_22") = DiasPago(Fecha)
        dr("acu_22") = 0
        dr("his_22") = 0

        'SECCION COMERCIAL
        dr("dia_3") = Ventas(Fecha, Fecha)
        dr("acu_3") = Ventas(New Date(Fecha.Year, Fecha.Month, 1), Fecha)
        dr("his_3") = Ventas(f1, f2)

        dr("dia_4") = VentasFamilia(Fecha, Fecha, "301", "501")
        dr("acu_4") = VentasFamilia(New Date(Fecha.Year, Fecha.Month, 1), Fecha, "301", "501")
        dr("his_4") = VentasFamilia(f1, f2, "301", "501")

        dr("dia_5") = VentasFamilia(Fecha, Fecha, "301", "502")
        dr("acu_5") = VentasFamilia(New Date(Fecha.Year, Fecha.Month, 1), Fecha, "301", "502")
        dr("his_5") = VentasFamilia(f1, f2, "301", "502")

        dr("dia_7") = Sistemas(Fecha, Fecha)
        dr("acu_7") = Sistemas(New Date(Fecha.Year, Fecha.Month, 1), Fecha)
        dr("his_7") = Sistemas(f1, f2)

        dr("dia_20") = VentasFamilia(Fecha, Fecha, "304", "533")
        dr("acu_20") = VentasFamilia(New Date(Fecha.Year, Fecha.Month, 1), Fecha, "304", "533")
        dr("his_20") = VentasFamilia(f1, f2, "304", "533")

        'SERVICES
        dr("dia_6") = Service(Fecha, Fecha)
        dr("acu_6") = Service(New Date(Fecha.Year, Fecha.Month, 1), Fecha)
        dr("his_6") = Service(f1, f2)

        dr("dia_9") = DiasService()
        dr("acu_9") = 0
        dr("his_9") = 0

        dr("dia_21") = DiasServicePromedio()
        dr("acu_21") = 0
        dr("his_21") = 0

        dr("dia_10") = EquiposPlanta()
        dr("acu_10") = 0
        dr("his_10") = 0

        dr("dia_24") = CantidadRetirarPendientesinruta(New Date(Fecha.Year, Fecha.Month, 1), New Date(Fecha.Year, Fecha.Month, 1).AddMonths(1))
        dr("acu_24") = 0
        dr("his_24") = 0

        dr("dia_17") = CantidadRetirarPendiente(New Date(Fecha.Year, Fecha.Month, 1), New Date(Fecha.Year, Fecha.Month, 1).AddMonths(1))
        dr("acu_17") = 0
        dr("his_17") = 0

        dr("dia_18") = Cantidadretirar(New Date(Fecha.Year, Fecha.Month, 1), New Date(Fecha.Year, Fecha.Month, 1).AddMonths(1))
        dr("acu_18") = 0
        dr("his_18") = 0

        dr("dia_19") = VenceMes(New Date(Fecha.Year, Fecha.Month, 1), New Date(Fecha.Year, Fecha.Month, 1).AddMonths(1))
        dr("acu_19") = 0
        dr("his_19") = 0

        dr("dia_25") = cantidadretirarpromedio(New Date(Fecha.Year, Fecha.Month, 1), New Date(Fecha.Year, Fecha.Month, 1).AddMonths(1))
        dr("acu_25") = 0
        dr("his_25") = cantidadretirarpromedio(f1, f2)

saltar:

        'PRODUCCION
        dr("dia_8") = DiasStock(New Date(Fecha.Year - 1, Fecha.Month, Fecha.Day), Fecha)
        dr("acu_8") = 0
        dr("his_8") = 0

        dr("dia_11") = 0
        dr("acu_11") = PromedioAbonos(New Date(Fecha.Year, Fecha.Month, 1), Fecha, 3)
        dr("his_11") = PromedioAbonos(f1, f2, 3)

        dr("dia_12") = 0
        dr("acu_12") = PromedioAbonos(New Date(Fecha.Year, Fecha.Month, 1), Fecha, 2)
        dr("his_12") = PromedioAbonos(f1, f2, 2)

        dr("dia_13") = 0
        dr("acu_13") = PromedioEquiposDomicilios(New Date(Fecha.Year, Fecha.Month, 1), Fecha, 3)
        dr("his_13") = PromedioEquiposDomicilios(f1, f2, 3)

        dr("dia_14") = 0
        dr("acu_14") = PromedioEquiposDomicilios(New Date(Fecha.Year, Fecha.Month, 1), Fecha, 2)
        dr("his_14") = PromedioEquiposDomicilios(f1, f2, 2)

        dr("dia_15") = 0
        dr("acu_15") = PromedioEquiposAcompanantes(New Date(Fecha.Year, Fecha.Month, 1), Fecha, 3)
        dr("his_15") = PromedioEquiposAcompanantes(f1, f2, 3)

        dr("dia_16") = 0
        dr("acu_16") = PromedioEquiposAcompanantes(New Date(Fecha.Year, Fecha.Month, 1), Fecha, 2)
        dr("his_16") = PromedioEquiposAcompanantes(f1, f2, 2)

        'ausencias
        dr("dia_23") = Ausencias(Fecha, Fecha)
        dr("acu_23") = Ausencias(New Date(Fecha.Year, Fecha.Month, 1), Fecha)
        dr("his_23") = 0

        dr.EndEdit()
    End Sub
    Public Sub Grabar()
        da.Update(dt)
    End Sub
    Public Sub EnviarMail()
        Dim Archivo As String
        Dim eMail As New MailMessage

        Archivo = GenerarReporte()

        With eMail
            .Subject = Archivo
            .From = New MailAddress("noreply@matafuegosgeorgia.com", "Sistema Automatico")
            Destinatarios(eMail)
            .Attachments.Add(New Attachment(Archivo))
        End With

        Try
            'Smtp.Send(eMail)
            eMail.Dispose()

        Catch ex As Exception

        End Try

    End Sub
    Public Function GenerarReporte() As String
        Dim rpt As New ReportDocument
        Dim Archivo As String

        Archivo = "Indices " & Fecha.ToString("dd-MM-yyyy") & ".pdf"

        rpt.Load(RPTX3 & "\XTABLERO.rpt")
        rpt.SetDatabaseLogon(DB_USR, DB_PWD)

        rpt.SetParameterValue("dat", Fecha)

        rpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, Archivo)

        Return Archivo

    End Function
    Private Sub Destinatarios(ByVal eMail As MailMessage)
        Dim Mails As New StreamReader("mails.txt")
        Dim Linea As String

        Try
            Do
                Linea = Mails.ReadLine()

                If Not Linea Is Nothing Then
                    If ValidarMail(Linea) Then eMail.To.Add(Linea)
                End If

            Loop Until Linea Is Nothing

            Mails.Close()

        Catch ex As Exception

        End Try

    End Sub
    Private Function ValidarMail(ByVal sMail As String) As Boolean
        Return Regex.IsMatch(sMail, "^([\w-]+\.)*?[\w-]+@[\w-]+\.([\w-]+\.)*?[\w]+$")
    End Function
    'Administración y Finanzas
    Private Function Cobranzas(ByVal Desde As Date, ByVal Hasta As Date) As Double
        'Calcula las cobranzas realizadas en el dia, acumulado mensual
        'proyectado e historico
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim dt1 As New DataTable
        Dim tot As Double = 0
        Dim sih As New Factura(cn)
        Dim sns As Integer = 1

        'Armado de SQL
        Sql = "SELECT pth.num_0, dat_0, vcrtyp_0, vcrnum_0, snsdud_0, amtloc_0, snsdud_0 * amtloc_0 AS importe "
        Sql &= "FROM (paypth pth INNER JOIN paypthdoc ptdoc ON (pth.num_0 = ptdoc.num_0)) INNER JOIN bpcustomer bpc ON (pth.bpr_0 = bpcnum_0) "
        Sql &= "WHERE pth.sta_0 = 3 AND pth.typ_0 = 2 AND pth.dat_0 >= :datini_0 AND pth.dat_0 <= :datfin_0 "
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("datini_0", OracleType.DateTime).Value = Desde
        da.SelectCommand.Parameters.Add("datfin_0", OracleType.DateTime).Value = Hasta
        da.Fill(dt)

        For Each dr In dt.Rows
            sns = 1
            If dr("vcrtyp_0").ToString = "NDC" AndAlso Desde <> Hasta Then
                If sih.Abrir(dr("vcrnum_0").ToString) Then
                    'Consulto si el cbte contiene articulo de cheque rech.
                    If sih.ExisteArticulo("900068") OrElse sih.ExisteArticulo("900097") Then
                        sns = -1
                    End If
                End If
            End If

            tot += CDbl(dr("importe")) * sns
        Next

        Return tot

    End Function
    Private Function DiasCalle() As Integer
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim Deuda As Double = 0
        Dim Ventas As Double = 0

        'Armado de SQL - CALCULO DE TOTAL DE DEUDA
        Sql = "SELECT SUM((amtloc_0 - payloc_0) * sns_0) AS total "
        Sql &= "FROM gaccdudate "
        Sql &= "WHERE dudsta_0 = 2 AND flgcle_0 <> 2 AND sac_0 IN ('DVL', 'DVE', 'DGJ')"
        da = New OracleDataAdapter(Sql, cn)
        da.Fill(dt)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            If Not IsDBNull(dr("total")) Then
                Deuda = CDbl(dr("total"))
            End If

        End If

        'Armado de Sql - Ventas último año (365 dias hacia atrás)
        Sql = "SELECT SUM(ratcur_0 * amtati_0 * sns_0) AS total "
        Sql &= "FROM sinvoice "
        Sql &= "WHERE accdat_0 > :accdat_0 AND invtyp_0 <> 5"

        Sql = "SELECT SUM(ratcur_0 * sid.amtatilin_0 * sih.sns_0) AS total "
        Sql &= "FROM sinvoice sih INNER JOIN sinvoiced sid ON (sih.num_0 = sid.num_0) "
        Sql &= "WHERE accdat_0 > :accdat_0 AND invtyp_0 <> 5 AND itmref_0 NOT IN ('900106', '900097', '900068')"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("accdat_0", OracleType.DateTime).Value = Today.AddYears(-1)
        dt = New DataTable
        da.Fill(dt)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            If Not IsDBNull(dr("total")) Then
                Ventas = CDbl(dr("total"))
            End If

        End If

        da.Dispose()
        dt.Dispose()

        Return CInt(Deuda * 365 / Ventas)

    End Function
    Private Function DiasCobro(ByVal Fecha As Date) As Integer
        Dim da As OracleDataAdapter
        Dim Sql As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim FacturasDias As Double = 0
        Dim CobranzaDias As Double = 0
        Dim Acumulado As Double = 0

        'PROCESO DE FACTURAS
        Sql = "SELECT sih.accdat_0, amtati_0 * sns_0 AS importe "
        Sql &= "FROM (paypth pth INNER JOIN paypthdoc ptd ON (pth.num_0 = ptd.num_0)) INNER JOIN sinvoice sih ON (ptd.vcrnum_0 = sih.num_0) "
        Sql &= "WHERE dat_0 = :dat_0 AND pth.typ_0 = 2 "
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("dat_0", OracleType.DateTime).Value = Fecha
        dt = New DataTable
        da.Fill(dt)
        da.Dispose()

        For Each dr In dt.Rows
            Dim f As Date = CDate(dr("accdat_0"))
            Dim Dias As Long = DateDiff(DateInterval.Day, f, Fecha)
            Dim Importe As Double = CDbl(dr("importe"))

            Acumulado += Importe
            FacturasDias += Importe * Dias
        Next
        If Acumulado > 0 Then
            FacturasDias = CInt(FacturasDias / Acumulado)
        Else
            FacturasDias = 0
        End If


        'PROCESO DE RECIBOS - CHEQUES Y EFECTIVO
        Sql = "SELECT ptd.* "
        Sql &= "FROM paypth pth INNER JOIN payptd ptd ON (pth.num_0 = ptd.num_0) "
        Sql &= "WHERE pth.typ_0 = 2 AND dat_0 = :dat_0 "
        Sql &= "ORDER BY ptd.num_0, ptd.lig_0"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("dat_0", OracleType.DateTime).Value = Fecha
        dt = New DataTable
        da.Fill(dt)
        da.Dispose()

        Acumulado = 0

        For Each dr In dt.Rows
            Dim f As Date
            Dim Dias As Long
            Dim Importe As Double

            Select Case dr("pam_0").ToString
                Case "CON"
                    Dias = 0

                Case "CHT", "CHD", "CHQ"
                    f = CDate(dr("duddat_0"))

                    If f > Fecha Then
                        Dias = DateDiff(DateInterval.Day, Fecha, f)
                    Else
                        Dias = 0
                    End If


            End Select

            Importe = CDbl(dr("amtlin2_0"))
            Acumulado += Importe
            CobranzaDias += Importe * Dias

        Next
        'PROCESO DE RECIBOS - RETENCIONES
        Sql = "SELECT ptd.amtrtz_0 "
        Sql &= "FROM paypth pth INNER JOIN paypthrit ptd ON (pth.num_0 = ptd.num_0) "
        Sql &= "WHERE pth.typ_0 = 2 AND dat_0 = :dat_0 "
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("dat_0", OracleType.DateTime).Value = Fecha
        dt = New DataTable
        da.Fill(dt)
        da.Dispose()

        For Each dr In dt.Rows
            Dim Importe As Double

            Importe = CDbl(dr("amtrtz_0"))
            Acumulado += Importe

        Next
        If Acumulado > 0.5 Then
            CobranzaDias = CInt(CobranzaDias / Acumulado)
        Else
            CobranzaDias = 0
        End If

        Return CInt(FacturasDias + CobranzaDias)

    End Function
    Private Function DiasPago(ByVal Fecha As Date) As Integer
        Dim da As OracleDataAdapter
        Dim Sql As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim FacturasDias As Double = 0
        Dim CobranzaDias As Double = 0
        Dim Acumulado As Double = 0

        'PROCESO DE FACTURAS
        Sql = "SELECT sih.accdat_0, amtati_0 * sns_0 AS importe "
        Sql &= "FROM (paypth pth INNER JOIN paypthdoc ptd ON (pth.num_0 = ptd.num_0)) INNER JOIN pinvoice sih ON (ptd.vcrnum_0 = sih.num_0) "
        Sql &= "WHERE dat_0 = :dat_0 AND pth.typ_0 = 1 "
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("dat_0", OracleType.DateTime).Value = Fecha
        dt = New DataTable
        da.Fill(dt)
        da.Dispose()

        For Each dr In dt.Rows
            Dim f As Date = CDate(dr("accdat_0"))
            Dim Dias As Long = DateDiff(DateInterval.Day, f, Fecha)
            Dim Importe As Double = CDbl(dr("importe"))

            Acumulado += Importe
            FacturasDias += Importe * Dias
        Next
        If Acumulado > 0 Then
            FacturasDias = CInt(FacturasDias / Acumulado)
        Else
            FacturasDias = 0
        End If


        'PROCESO DE RECIBOS - CHEQUES Y EFECTIVO
        Sql = "SELECT ptd.* "
        Sql &= "FROM paypth pth INNER JOIN payptd ptd ON (pth.num_0 = ptd.num_0) "
        Sql &= "WHERE pth.typ_0 = 1 AND dat_0 = :dat_0 "
        Sql &= "ORDER BY ptd.num_0, ptd.lig_0"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("dat_0", OracleType.DateTime).Value = Fecha
        dt = New DataTable
        da.Fill(dt)
        da.Dispose()

        Acumulado = 0

        For Each dr In dt.Rows
            Dim f As Date
            Dim Dias As Long
            Dim Importe As Double

            Select Case dr("pam_0").ToString
                Case "CON"
                    Dias = 0

                Case "CHT", "CHD", "CHQ"
                    f = CDate(dr("duddat_0"))

                    If f > Fecha Then
                        Dias = DateDiff(DateInterval.Day, Fecha, f)
                    Else
                        Dias = 0
                    End If


            End Select

            Importe = CDbl(dr("amtlin2_0"))
            Acumulado += Importe
            CobranzaDias += Importe * Dias

        Next
        'PROCESO DE RECIBOS - RETENCIONES
        Sql = "SELECT ptd.amtrtz_0 "
        Sql &= "FROM paypth pth INNER JOIN paypthrit ptd ON (pth.num_0 = ptd.num_0) "
        Sql &= "WHERE pth.typ_0 = 2 AND dat_0 = :dat_0 "
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("dat_0", OracleType.DateTime).Value = Fecha
        dt = New DataTable
        da.Fill(dt)
        da.Dispose()

        For Each dr In dt.Rows
            Dim Importe As Double

            Importe = CDbl(dr("amtrtz_0"))
            Acumulado += Importe

        Next
        If Acumulado > 0.5 Then
            CobranzaDias = CInt(CobranzaDias / Acumulado)
        Else
            CobranzaDias = 0
        End If

        Return CInt(FacturasDias + CobranzaDias)

    End Function
    'Comercial
    Private Function Ventas(ByVal Desde As Date, ByVal Hasta As Date) As Double
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim dr As DataRow

        'Armado de SQL
        Sql = "SELECT SUM(amtnot_0 * sns_0) FROM sinvoice WHERE accdat_0 >= :datini_0 AND accdat_0 <= :datfin_0"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("datini_0", OracleType.DateTime).Value = Desde
        da.SelectCommand.Parameters.Add("datfin_0", OracleType.DateTime).Value = Hasta
        da.Fill(dt)

        If dt.Rows.Count = 0 Then
            Return 0

        Else
            dr = dt.Rows(0)
            If IsDBNull(dr(0)) Then
                Return 0

            Else
                Return CDbl(dr(0))
            End If

        End If

    End Function
    Private Function VentasFamilia(ByVal Desde As Date, ByVal Hasta As Date, ByVal codigo As String, ByVal Familia As String) As Integer
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim dr As DataRow

        'Armado de SQL
        Sql = "SELECT SUM(sid.qty_0 * sih.sns_0) "
        Sql &= "FROM (sinvoice sih INNER JOIN sinvoiced sid ON (sih.num_0 = sid.num_0)) INNER JOIN itmmaster itm on (sid.itmref_0 = itm.itmref_0) "
        Sql &= "WHERE tsicod_3 = :tsicod_3 AND tsicod_4 = :tsicod_4 AND accdat_0 >= :datini_0 AND accdat_0 <= :datfin_0"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("datini_0", OracleType.DateTime).Value = Desde
        da.SelectCommand.Parameters.Add("datfin_0", OracleType.DateTime).Value = Hasta
        da.SelectCommand.Parameters.Add(":tsicod_3", OracleType.VarChar).Value = codigo
        da.SelectCommand.Parameters.Add("tsicod_4", OracleType.VarChar).Value = Familia
        da.Fill(dt)

        If dt.Rows.Count > 0 Then

            dr = dt.Rows(0)
            If IsDBNull(dr(0)) Then
                Return 0

            Else
                Return CInt(dr(0))

            End If

        Else
            Return 0

        End If

    End Function
    Private Function Sistemas(ByVal Desde As Date, ByVal Hasta As Date) As Integer
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim dr As DataRow

        'Armado de SQL
        Sql = "SELECT SUM(sid.qty_0 * sih.sns_0) "
        Sql &= "FROM (sinvoice sih INNER JOIN sinvoiced sid ON (sih.num_0 = sid.num_0)) INNER JOIN itmmaster itm on (sid.itmref_0 = itm.itmref_0) "
        Sql &= "WHERE tsicod_3 = '304' AND tsicod_4 IN ('531', '532') AND accdat_0 >= :datini_0 AND accdat_0 <= :datfin_0"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("datini_0", OracleType.DateTime).Value = Desde
        da.SelectCommand.Parameters.Add("datfin_0", OracleType.DateTime).Value = Hasta
        da.Fill(dt)

        If dt.Rows.Count > 0 Then

            dr = dt.Rows(0)
            If IsDBNull(dr(0)) Then
                Return 0

            Else
                Return CInt(dr(0))

            End If

        Else
            Return 0

        End If
    End Function
    'Services
    Private Function Service(ByVal Desde As Date, ByVal Hasta As Date) As Integer
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim dr As DataRow

        'Armado de SQL
        Sql = "select count(sre.srenum_0) as cantidad from sremac sre inner join serrequest ser on (sre.srenum_0 = ser.srenum_0) where sre.creusr_0 = 'RECEP' "
        Sql &= "and sre.credat_0 between :fechaini_0 and :fechafin_0 and ser.SREBPC_0 <> '402000'"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("fechaini_0", OracleType.DateTime).Value = Desde
        da.SelectCommand.Parameters.Add("fechafin_0", OracleType.DateTime).Value = Hasta
        da.Fill(dt)
        dr = dt.Rows(0)
        Return CInt(dr("cantidad"))

    End Function
    Private Function DiasService() As Double
        Dim da As OracleDataAdapter
        Dim Sql As String
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim TotalEquipos As Integer = 0
        Dim Total As Integer

        Sql = "SELECT * "
        Sql &= "FROM xsegto2 xse INNER JOIN interven itn ON (xse.itn_0 = itn.num_0) "
        Sql &= "WHERE dat_4 = to_date('31/12/1599', 'dd/mm/yyyy') AND dat_3 > xse.dat_0 AND zflgtrip_0 <> 8 and itn.bpc_0 <> '402000' "

        da = New OracleDataAdapter(Sql, cn)
        da.Fill(dt)
        da.Dispose()

        For Each dr In dt.Rows
            Dim FechaIngreso As Date = CDate(dr("dat_0"))
            Dim FechaProcesado As Date = CDate(dr("dat_3"))
            Dim Dias As Integer = CInt(DateDiff(DateInterval.Day, FechaIngreso, FechaProcesado))
            Dim Equipos As Integer = 0

            'Sumo equipos
            Equipos += CInt(dr("cant_0"))
            Equipos += CInt(dr("cant_1"))
            Equipos += CInt(dr("rech_0"))
            Equipos += CInt(dr("rech_1"))

            Total += Dias * Equipos
            TotalEquipos += Equipos

        Next

        If TotalEquipos > 0 Then
            Return Total / TotalEquipos
        Else
            Return 0
        End If

    End Function
    Private Function DiasServicePromedio() As Double
        Dim da As OracleDataAdapter
        Dim Sql As String
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim TotalEquipos As Integer = 0
        Dim Total As Integer

        Sql = "SELECT * "
        Sql &= "FROM xsegto2 xse INNER JOIN interven itn ON (xse.itn_0 = itn.num_0) "
        Sql &= "WHERE dat_2 <> to_date('31/12/1599', 'dd/mm/yyyy') AND dat_2 > xse.dat_0 AND zflgtrip_0 <> 8 and itn.bpc_0 <> '402000'"

        da = New OracleDataAdapter(Sql, cn)
        da.Fill(dt)
        da.Dispose()

        For Each dr In dt.Rows
            Dim FechaIngreso As Date = CDate(dr("dat_0"))
            Dim FechaProcesado As Date = CDate(dr("dat_2"))
            Dim Dias As Integer = CInt(DateDiff(DateInterval.Day, FechaIngreso, FechaProcesado))
            Dim Equipos As Integer = 0

            'Sumo equipos
            Equipos += CInt(dr("cant_0"))
            Equipos += CInt(dr("cant_1"))
            Equipos += CInt(dr("rech_0"))
            Equipos += CInt(dr("rech_1"))

            Total += Dias * Equipos
            TotalEquipos += Equipos

        Next

        If TotalEquipos > 0 Then
            Return Total / TotalEquipos
        Else
            Return 0
        End If

    End Function
    Private Function EquiposPlanta() As Integer
        Dim da As OracleDataAdapter
        Dim Sql As String
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim Total As Integer = 0

        Sql = "SELECT xse.* "
        Sql &= "FROM xsegto2 xse INNER JOIN interven itn ON (itn_0 = num_0) "
        Sql &= "WHERE dat_4 = to_date('31/12/1599', 'dd/mm/yyyy') AND xsector_0 <> 'ACH' AND zflgtrip_0 IN (2,3,4,6,7) and itn.bpc_0 <> '402000' "

        da = New OracleDataAdapter(Sql, cn)
        da.Fill(dt)
        da.Dispose()

        For Each dr In dt.Rows
            If CInt(dr("cant_0")) = 0 And CInt(dr("cant_1")) = 0 And CInt(dr("rech_0")) = 0 And CInt(dr("rech_1")) = 0 Then
                Total += CInt(dr("equipos_0"))

            Else
                Total += CInt(dr("cant_0"))
                Total += CInt(dr("cant_1"))
                Total += CInt(dr("rech_0"))
                Total += CInt(dr("rech_1"))

            End If
        Next

        Return Total

    End Function
    Private Function CantidadRetirarPendiente(ByVal fecha_inicio As Date, ByVal fecha_fin As Date) As Double
        Dim da As OracleDataAdapter
        Dim Sql As String
        Dim dt As New DataTable
        Dim dr As DataRow

        Sql = " Select SUM(tqty_0) as cantidad "
        Sql &= "FROM (interven itn INNER JOIN yitndet yit ON (itn.num_0 = yit.num_0)) INNER JOIN itmmaster itm ON (yit.itmref_0 = itm.itmref_0) "
        Sql &= "WHERE  (itn.dat_0 >= :fechainicio_0 and itn.dat_0 < :fechafin_0) AND "
        Sql &= "typlig_0 = 1 and zflgtrip_0 = 1  and "
        Sql &= "cfglin_0 IN ('451', '505') and itn.bpc_0 <> '402000'"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("fechainicio_0", OracleType.DateTime).Value = fecha_inicio
        da.SelectCommand.Parameters.Add("fechafin_0", OracleType.DateTime).Value = fecha_fin
        da.Fill(dt)
        da.Dispose()
        dr = dt.Rows(0)

        Return CDbl(dr("cantidad"))
    End Function
    Private Function CantidadRetirarPendientesinruta(ByVal fecha_inicio As Date, ByVal fecha_fin As Date) As Double
        Dim da As OracleDataAdapter
        Dim Sql As String
        Dim dt As New DataTable
        Dim dr As DataRow
        Sql = " Select SUM(tqty_0) as cantidad "
        Sql &= "FROM (interven itn INNER JOIN yitndet yit ON (itn.num_0 = yit.num_0)) INNER JOIN itmmaster itm ON (yit.itmref_0 = itm.itmref_0) "
        Sql &= "WHERE  (itn.dat_0 >= :fechainicio_0 and itn.dat_0 < :fechafin_0) AND "
        Sql &= "typlig_0 = 1 and zflgtrip_0 = 1 and tripnum_0 = ' ' and "
        Sql &= "cfglin_0 IN ('451', '505') and itn.bpc_0 <> '402000' "
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("fechainicio_0", OracleType.DateTime).Value = fecha_inicio
        da.SelectCommand.Parameters.Add("fechafin_0", OracleType.DateTime).Value = fecha_fin
        da.Fill(dt)
        da.Dispose()
        dr = dt.Rows(0)
        Return CDbl(dr("cantidad"))
    End Function
    Private Function Cantidadretirar(ByVal fecha_inicio As Date, ByVal fecha_fin As Date) As Double
        Dim da As OracleDataAdapter
        Dim Sql As String
        Dim dt As New DataTable
        Dim dr As DataRow
        Sql = " Select SUM(tqty_0) as cantidad "
        Sql &= "FROM (interven itn INNER JOIN yitndet yit ON (itn.num_0 = yit.num_0)) INNER JOIN itmmaster itm ON (yit.itmref_0 = itm.itmref_0) "
        Sql &= "WHERE  (itn.dat_0 >= :fechainicio_0 and itn.dat_0 < :fechafin_0) AND "
        Sql &= "typlig_0 = 1 and "
        Sql &= "cfglin_0 IN ('451', '505') and itn.bpc_0 <> '402000'"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("fechainicio_0", OracleType.DateTime).Value = fecha_inicio
        da.SelectCommand.Parameters.Add("fechafin_0", OracleType.DateTime).Value = fecha_fin
        da.Fill(dt)
        da.Dispose()
        dr = dt.Rows(0)
        Return CDbl(dr("cantidad"))

    End Function
    Private Function VenceMes(ByVal fecha_inicio As Date, ByVal fecha_fin As Date) As Double
        Dim da As OracleDataAdapter
        Dim Sql As String
        Dim dt As New DataTable
        Dim dr As DataRow
        Sql = "Select count(mac.macnum_0) as cantidad "
        Sql &= "FROM (machines mac INNER JOIN ymacitm ymc ON (mac.macnum_0 = ymc.macnum_0)) INNER JOIN bomd bom ON (mac.macpdtcod_0 = bom.itmref_0 AND ymc.cpnitmref_0 = bom.cpnitmref_0) "
        Sql &= "WHERE bomalt_0 = 99 AND bomseq_0 = '10' AND (datnext_0 >= :fechainicio_0 and datnext_0 < :fechafin_0) and mac.bpcnum_0 <> '402000' AND xitn_0 = ' ' "
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("fechainicio_0", OracleType.DateTime).Value = fecha_inicio
        da.SelectCommand.Parameters.Add("fechafin_0", OracleType.DateTime).Value = fecha_fin
        da.Fill(dt)
        da.Dispose()
        dr = dt.Rows(0)
        Return CDbl(dr("cantidad"))
    End Function
    Private Function cantidadretirarpromedio(ByVal fecha_inicio As Date, ByVal fecha_fin As Date) As Double
        Dim da As OracleDataAdapter
        Dim Sql As String
        Dim dt As New DataTable
        Dim dr As DataRow
        Sql = " Select SUM(tqty_0) as cantidad, count(itn.num_0)as  promedio "
        Sql &= "FROM (interven itn INNER JOIN yitndet yit ON (itn.num_0 = yit.num_0)) INNER JOIN itmmaster itm ON (yit.itmref_0 = itm.itmref_0) "
        Sql &= "WHERE  (itn.dat_0 >= :fechainicio_0 and itn.dat_0 < :fechafin_0) AND "
        Sql &= "typlig_0 = 1 and "
        Sql &= "cfglin_0 IN ('451', '505') and itn.bpc_0 <> '402000'"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("fechainicio_0", OracleType.DateTime).Value = fecha_inicio
        da.SelectCommand.Parameters.Add("fechafin_0", OracleType.DateTime).Value = fecha_fin
        da.Fill(dt)
        da.Dispose()
        dr = dt.Rows(0)
        Return (CInt(dr("cantidad")) / CInt(dr("promedio")))

    End Function
    'Produccion
    Private Function DiasStock(ByVal desde As Date, ByVal hasta As Date) As Double
        Dim da As OracleDataAdapter
        Dim dt As DataTable
        Dim Sql As String
        Dim dr As DataRow
        Dim ultimo As Integer
        Dim ProDiario As Integer

        Sql = "select sum(itmc.vlttot_0*itm.physto_0) as suma "
        Sql &= "from itmcost itmc inner join itmmvt itm on (itmc.itmref_0 = itm.itmref_0) and (itmc.stofcy_0 = itm.stofcy_0) "
        Sql &= "where  itmc.yea_0 = '2015' and itmc.csttyp_0 = '1' and itmc.csttot_0 <> '0' "
        Sql &= "order by itmc.itcdat_0 desc"
        da = New OracleDataAdapter(Sql, cn)
        dt = New DataTable
        da.Fill(dt)
        da.Dispose()
        dr = dt.Rows(0)
        ultimo = CInt(dr("suma"))
        Sql = "select ((sum(amtati_0*sns_0))) as pro  from sinvoice where (accdat_0 between :fechaini and :fechafin) and (sivtyp_0 <> 'PRF') "
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("fechaini", OracleType.DateTime).Value = desde
        da.SelectCommand.Parameters.Add("fechafin", OracleType.DateTime).Value = hasta
        dt = New DataTable
        da.Fill(dt)
        da.Dispose()
        dr = dt.Rows(0)
        ProDiario = CInt(dr("pro"))

        Return ultimo / (ProDiario / 365)

    End Function
    Private Function PromedioAbonos(ByVal Fecha_inicio As Date, ByVal fecha_fin As Date, ByVal sector As Integer) As Double
        Dim da As OracleDataAdapter
        Dim dt As DataTable
        Dim Sql As String
        Dim Domicilios As Integer
        Dim Camionetas As Integer
        Dim Promedio As Double = 0
        Dim dr As DataRow

        '***********************************************************
        'CALCULO DE LA CANTIDAD DE DOMICILIOS VISITADOS
        '***********************************************************
        Sql = "SELECT count(xd.vcrnum_0) "
        Sql &= "FROM (xrutac xc INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0)) "
        Sql &= "INNER JOIN zunitrans tr ON (xc.patente_0 = tr.patnum_0) and (xc.transporte_0 = tr.bptnum_0) "
        Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND "
        Sql &= "      xc.fecha_0 <= :fechafin_0 AND "
        Sql &= "      tr.xsector_0 = :sector_0 AND "
        Sql &= "      tipo_0 in ('RET','ENT','NCI','NUE') --AND "
        'Sql &= "      xd.estado_0 = 3 AND "
        'Sql &= "      xc.valid_0 = 1 "

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("fechaini_0", OracleType.DateTime).Value = Fecha_inicio
        da.SelectCommand.Parameters.Add("fechafin_0", OracleType.DateTime).Value = fecha_fin
        da.SelectCommand.Parameters.Add("sector_0", OracleType.Number).Value = sector
        dt = New DataTable
        da.Fill(dt)
        da.Dispose()

        dr = dt.Rows(0)

        If IsDBNull(dr(0)) Then
            Domicilios = 0

        Else
            Domicilios = CInt(dr(0))

        End If

        '***********************************************************
        'CALCULO DE LA CANTIDAD DE CAMIONETAS UTILIZADAS
        '***********************************************************
        Sql = "SELECT DISTINCT fecha_0, transporte_0, patente_0 "
        Sql &= "FROM (xrutac xc INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0)) "
        Sql &= "INNER JOIN zunitrans tr ON (xc.patente_0 = tr.patnum_0) and (xc.transporte_0 = tr.bptnum_0) "
        Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND "
        Sql &= "      xc.fecha_0 <= :fechafin_0 AND "
        Sql &= "      tr.xsector_0 = :sector_0 AND "
        Sql &= "      tipo_0 in ('RET','ENT','NCI','NUE') AND "
        Sql &= "      xd.estado_0 = 3 AND "
        Sql &= "      xc.valid_0 = 1 "

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("fechaini_0", OracleType.DateTime).Value = Fecha_inicio
        da.SelectCommand.Parameters.Add("fechafin_0", OracleType.DateTime).Value = fecha_fin
        da.SelectCommand.Parameters.Add("sector_0", OracleType.Number).Value = sector
        dt = New DataTable
        da.Fill(dt)
        da.Dispose()

        Camionetas = dt.Rows.Count

        If Camionetas > 0 Then
            Promedio = Domicilios / Camionetas
        End If

        Return Promedio

    End Function
    Private Function PromedioEquiposDomicilios(ByVal FechaInicio As Date, ByVal FechaFin As Date, ByVal sector As Integer) As Double
        Dim da As OracleDataAdapter
        Dim Sql As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim equipos As Double
        Dim domicilio As Double
        Dim promedio As Double
        Dim Camionetas As Integer

        Sql = "SELECT SUM(equipos_1 + equipos_3)as equipos "
        Sql &= "FROM xrutac xc INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0)"
        Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
        Sql &= "WHERE fecha_0 BETWEEN :FechaInicio AND :FechaFin and tr.xsector_0 = :sector_0 and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' "


        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("FechaInicio", OracleType.DateTime).Value = FechaInicio
        da.SelectCommand.Parameters.Add("FechaFin", OracleType.DateTime).Value = FechaFin
        da.SelectCommand.Parameters.Add("sector_0", OracleType.Number).Value = sector
        dt = New DataTable
        da.Fill(dt)
        da.Dispose()
        dr = dt.Rows(0)
        If IsDBNull(dr("equipos")) Then Exit Function
        equipos = CDbl(dr("equipos"))

        Sql = " SELECT sum(count(xd.vcrnum_0))AS DOMI FROM xrutac xc "
        Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
        Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
        Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 and tipo_0 in ('RET','ENT','NCI','NUE') "
        Sql &= "AND xd.estado_0 = '3' and xc.valid_0 = '1' "
        Sql &= "group by xc.patente_0"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("fechaini_0", OracleType.DateTime).Value = FechaInicio
        da.SelectCommand.Parameters.Add("fechafin_0", OracleType.DateTime).Value = FechaFin
        da.SelectCommand.Parameters.Add("sector_0", OracleType.Number).Value = sector
        dt = New DataTable
        da.Fill(dt)
        da.Dispose()
        dr = dt.Rows(0)
        domicilio = CInt(dr("DOMI"))
        'domicilio = CDbl(dr("domicilio"))
        promedio = equipos / domicilio

        If sector = 3 Then
            Sql = " SELECT DISTINCT(SELECT COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            Sql &= "FROM xrutac xc "
            Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'BQD436' ) "
            Sql &= "+ "
            Sql &= "(SELECT COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            Sql &= "FROM xrutac xc "
            Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0  "
            Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'DSE539') "
            Sql &= "+ "
            Sql &= "(SELECT  COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            Sql &= "FROM xrutac xc "
            Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'RBK551') "
            Sql &= "+ "
            Sql &= "(SELECT  COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            Sql &= "FROM xrutac xc "
            Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0  "
            Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'CNI308') "
            Sql &= "+ "
            Sql &= "(SELECT  COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            Sql &= "FROM xrutac xc "
            Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'LMX439') "
            Sql &= "+ "
            Sql &= "(SELECT  COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            Sql &= "FROM xrutac xc "
            Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'CUX818') "
            Sql &= "+ "
            Sql &= "(SELECT  COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            Sql &= "FROM xrutac xc "
            Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'KVI728') "
            Sql &= "+ "
            Sql &= "(SELECT  COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            Sql &= "FROM xrutac xc "
            Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'AYH290') "
            Sql &= "+ "
            Sql &= "(SELECT COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            Sql &= "FROM xrutac xc "
            Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'BBJ106') AS SUMA FROM XRUTAC "
        Else
            Sql = "SELECT DISTINCT(SELECT COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            Sql &= "FROM xrutac xc "
            Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'DZH386' ) "
            Sql &= "+ "
            Sql &= "(SELECT COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            Sql &= "FROM xrutac xc "
            Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'CNG703') "
            Sql &= "+ "
            Sql &= "(SELECT  COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            Sql &= "FROM xrutac xc "
            Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'CFC807') "
            Sql &= "+ "
            Sql &= "(SELECT  COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            Sql &= "FROM xrutac xc "
            Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'NBB532') "
            Sql &= "+ "
            Sql &= "(SELECT  COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            Sql &= "FROM xrutac xc "
            Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'ECU604') "
            Sql &= "+ "
            Sql &= "(SELECT COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            Sql &= "FROM xrutac xc "
            Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'CZT252') "
            Sql &= "+ "
            Sql &= "(SELECT COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            Sql &= "FROM xrutac xc "
            Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'FAH030') "
            Sql &= "+ "
            Sql &= "(SELECT COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            Sql &= "FROM xrutac xc "
            Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'CUY968') "
            Sql &= "+ "
            Sql &= "(SELECT COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            Sql &= "FROM xrutac xc "
            Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'CDJ453') "
            Sql &= "+ "
            Sql &= "(SELECT COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            Sql &= "FROM xrutac xc "
            Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'KPN936') "
            Sql &= "+ "
            Sql &= "(SELECT COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            Sql &= "FROM xrutac xc "
            Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'CMD866') "
            Sql &= "AS SUMA FROM XRUTAC "
            '    Sql = " SELECT DISTINCT(SELECT COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            '    Sql &= "FROM xrutac xc "
            '    Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            '    Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            '    Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            '    Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'GDO111' ) "
            '    Sql &= "+ "
            '    Sql &= "(SELECT COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            '    Sql &= "FROM xrutac xc "
            '    Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            '    Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            '    Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0  "
            '    Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'RBK551') "
            '    Sql &= "+ "
            '    Sql &= "(SELECT  COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            '    Sql &= "FROM xrutac xc "
            '    Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            '    Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            '    Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            '    Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'BQD436') "
            '    Sql &= "+ "
            '    Sql &= "(SELECT  COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            '    Sql &= "FROM xrutac xc "
            '    Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            '    Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            '    Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0  "
            '    Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'AYH290') "
            '    Sql &= "+ "
            '    Sql &= "(SELECT  COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            '    Sql &= "FROM xrutac xc "
            '    Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            '    Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            '    Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            '    Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'AJU817') "
            '    Sql &= "+ "
            '    Sql &= "(SELECT COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            '    Sql &= "FROM xrutac xc "
            '    Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            '    Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            '    Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            '    Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'DZW376') AS SUMA FROM XRUTAC "
            'Else

            '    Sql = "SELECT DISTINCT(SELECT COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            '    Sql &= "FROM xrutac xc "
            '    Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            '    Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            '    Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            '    Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'NBB532' ) "
            '    Sql &= "+ "
            '    Sql &= "(SELECT COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            '    Sql &= "FROM xrutac xc "
            '    Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            '    Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            '    Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            '    Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'CUY968') "
            '    Sql &= "+ "
            '    Sql &= "(SELECT  COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            '    Sql &= "FROM xrutac xc "
            '    Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            '    Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            '    Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            '    Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'ANG954') "
            '    Sql &= "+ "
            '    Sql &= "(SELECT  COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            '    Sql &= "FROM xrutac xc "
            '    Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            '    Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            '    Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            '    Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'FAH030') "
            '    Sql &= "+ "
            '    Sql &= "(SELECT  COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            '    Sql &= "FROM xrutac xc "
            '    Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            '    Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            '    Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            '    Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'DGL028') "
            '    Sql &= "+ "
            '    Sql &= "(SELECT COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            '    Sql &= "FROM xrutac xc "
            '    Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            '    Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            '    Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            '    Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'DZH386') "
            '    Sql &= "+ "
            '    Sql &= "(SELECT COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            '    Sql &= "FROM xrutac xc "
            '    Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            '    Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            '    Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            '    Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'KPN936') "
            '    Sql &= "+ "
            '    Sql &= "(SELECT COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            '    Sql &= "FROM xrutac xc "
            '    Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            '    Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            '    Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            '    Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = '807') "
            '    Sql &= "+ "
            '    Sql &= "(SELECT COUNT(distinct xc.FECHA_0) AS CANTIDAD "
            '    Sql &= "FROM xrutac xc "
            '    Sql &= "INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
            '    Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
            '    Sql &= "WHERE xc.fecha_0 >= :fechaini_0 AND xc.fecha_0 <= :fechafin_0 and tr.xsector_0 = :sector_0 "
            '    Sql &= "and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1' and xc.patente_0 = 'CZT252') "
            '    Sql &= "AS SUMA FROM XRUTAC "
        End If

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("fechaini_0", OracleType.DateTime).Value = FechaInicio
        da.SelectCommand.Parameters.Add("fechafin_0", OracleType.DateTime).Value = FechaFin
        da.SelectCommand.Parameters.Add("sector_0", OracleType.Number).Value = sector
        dt = New DataTable
        da.Fill(dt)
        da.Dispose()
        dr = dt.Rows(0)
        Camionetas = CInt(dr("SUMA"))


        Return promedio

        'Return CDbl(dr("promedio"))

        'For Each dr In dt.Rows
        '    Dim a, b As Integer

        '    a = CInt(dr("equipos"))
        '    b = CInt(dr("domicilios"))

        '    TotalAcompanante += a * b
        '    TotalEquipos += a
        'Next

        'If TotalEquipos > 0 Then
        '    Return TotalAcompanante / TotalEquipos
        'Else
        '    Return 0
        'End If


    End Function
    Private Function PromedioEquiposAcompanantes(ByVal FechaInicio As Date, ByVal FechaFin As Date, ByVal sector As Integer) As Double
        Dim da As OracleDataAdapter
        Dim Sql As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim TotalEquipos As Integer
        Dim TotalAcompanantes As Integer

        'Sql = "SELECT xc.ruta_0, acomp_0, acomp_1, acomp_2, SUM(equipos_1 + equipos_3) as equipos "
        'Sql &= "FROM xrutac xc INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
        'Sql &= "WHERE fecha_0 BETWEEN :FechaInicio AND :FechaFin AND tipo_0 IN ('RET', 'ENT') "
        'Sql &= "GROUP BY xc.ruta_0, acomp_0, acomp_1, acomp_2 "
        'Sql &= "ORDER BY ruta_0"
        Sql = " SELECT xc.ruta_0, acomp_0, acomp_1, acomp_2, SUM(equipos_1 + equipos_3) as equipos "
        Sql &= "FROM xrutac xc INNER JOIN xrutad xd ON (xc.ruta_0 = xd.ruta_0) "
        Sql &= "inner join zunitrans tr on (xc.PATENTE_0 = tr.patnum_0) and (xc.TRANSPORTE_0 = tr.bptnum_0) "
        Sql &= "WHERE fecha_0 BETWEEN :FechaInicio AND :FechaFin AND tr.xsector_0 = :sector_0 and tipo_0 in ('RET','ENT','NCI','NUE') AND xd.estado_0 = '3' and xc.valid_0 = '1'  "
        Sql &= "GROUP BY xc.ruta_0, acomp_0, acomp_1, acomp_2 having acomp_0 <> 0 or  acomp_1 <> 0 or acomp_2 <> 0 "
        Sql &= "ORDER BY ruta_0"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("FechaInicio", OracleType.DateTime).Value = FechaInicio
        da.SelectCommand.Parameters.Add("FechaFin", OracleType.DateTime).Value = FechaFin
        da.SelectCommand.Parameters.Add("sector_0", OracleType.Number).Value = sector
        dt = New DataTable
        da.Fill(dt)
        da.Dispose()

        For Each dr In dt.Rows
            Dim a, b As Integer

            a = CInt(dr("equipos"))

            b = 0
            If CInt(dr("acomp_0")) > 0 Then b += 1
            If CInt(dr("acomp_1")) > 0 Then b += 1
            If CInt(dr("acomp_2")) > 0 Then b += 1

            TotalAcompanantes += b
            TotalEquipos += a
        Next


        If TotalAcompanantes > 0 Then
            Return TotalEquipos / TotalAcompanantes

        Else
            Return 0
        End If


    End Function

    'Ausencias
    Private Function Ausencias(ByVal fecha_inicio As Date, ByVal fecha_fin As Date) As Double
        Dim dbConnection As OleDbConnection
        Dim dbDataTable As New Data.DataTable
        ' Dim dbDataSet As Data.DataSet
        Dim dbDataAdapter As OleDbDataAdapter
        Dim dr As DataRow
        Dim CadenaConexion As String
        Dim ausencia As Double
        Dim dotacion As Double
        '  Dim bdd As String

        CadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=\\servidor2\reloj\ausencias.mdb;"
        dbConnection = New OleDbConnection(CadenaConexion)
        dbDataAdapter = New OleDbDataAdapter("select * from ausencias where fecha >=? and fecha <=?;", dbConnection)
        dbDataAdapter.SelectCommand.Parameters.Add("fechainicio", OleDbType.Date)
        dbDataAdapter.SelectCommand.Parameters.Add("fechafin", OleDbType.Date)
        dbDataAdapter.SelectCommand.Parameters("fechainicio").Value = fecha_inicio
        dbDataAdapter.SelectCommand.Parameters("fechafin").Value = fecha_fin
        dbDataAdapter.Fill(dbDataTable)

        If dbDataTable.Rows.Count = 0 Then
            Return 0
        Else
            dr = dbDataTable.Rows(0)
            For Each dr In dbDataTable.Rows
                ausencia += (CDbl(dr("ausencia")))
                dotacion += (CDbl(dr("dotacion")))
            Next
            Return (ausencia / dotacion) * 100
        End If

       
        dbDataAdapter.Dispose()
    End Function
  
End Class