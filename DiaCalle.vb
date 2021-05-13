Imports System.Data.OracleClient

Public Class DiaCalle
    Private cn As OracleConnection
    Private usr As Usuario

    Public Sub New(ByVal cn As OracleConnection, ByVal Usr As Usuario)
        Me.cn = cn
        Me.usr = Usr
    End Sub
    Public Sub Calcular()

        Dim daVendedores As OracleDataAdapter
        Dim daDeuda As OracleDataAdapter
        Dim daVenta As OracleDataAdapter
        Dim Sql As String
        Dim dtVendedores As New DataTable
        Dim dtDeuda As New DataTable
        Dim dtVenta As New DataTable
        Dim tmp As New Temporal(cn, Me.usr, "DCALLE")

        tmp.Abrir()
        tmp.LimpiarTabla()
        tmp.Grabar()

        'Recupero todos los vendedores
        Sql = "SELECT repnum_0, repnam_0 from salesrep"
        daVendedores = New OracleDataAdapter(Sql, cn)
        daVendedores.Fill(dtVendedores)

        'Recupero toda la deuda
        Sql = "SELECT bpc.rep_0, sum((amtloc_0 - payloc_0) * sns_0) AS total "
        Sql &= "FROM gaccdudate gac inner join "
        Sql &= "     bpcustomer bpc on (gac.bpr_0 = bpc.bpcnum_0) "
        Sql &= "WHERE dudsta_0 = 2 AND "
        Sql &= "      flgcle_0 <> 2 AND "
        Sql &= "      sac_0 IN ('DVL', 'DVE', 'DGJ') "
        Sql &= "GROUP BY bpc.rep_0"
        daDeuda = New OracleDataAdapter(Sql, cn)
        daDeuda.Fill(dtDeuda)

        'Recupero la venta anual
        Sql = "SELECT bpc.rep_0, SUM(ratcur_0 * sid.amtatilin_0 * sih.sns_0) AS total "
        Sql &= "FROM sinvoice sih INNER JOIN "
        Sql &= "     sinvoiced sid ON (sih.num_0 = sid.num_0) inner join "
        Sql &= "     bpcustomer bpc on (sih.bpr_0 = bpc.bpcnum_0) "
        Sql &= "WHERE accdat_0 > :accdat_0 AND "
        Sql &= "      invtyp_0 <> 5 AND "
        Sql &= "      itmref_0 NOT IN ('900106', '900097', '900068') "
        Sql &= "group by bpc.rep_0"
        daVenta = New OracleDataAdapter(Sql, cn)
        daVenta.SelectCommand.Parameters.Add("accdat_0", OracleType.DateTime).Value = Today.AddYears(-1)
        daVenta.Fill(dtVenta)

        For Each drVendedor As DataRow In dtVendedores.Rows
            Dim dvDeuda As New DataView(dtDeuda)
            Dim dvVenta As New DataView(dtVenta)
            Dim rep As String

            'Filtro por cada vendedor
            rep = drVendedor("repnum_0").ToString
            dvDeuda.RowFilter = "rep_0 = '" & rep & "'"
            dvVenta.RowFilter = "rep_0 = '" & rep & "'"

            If dvDeuda.Count > 0 Or dvVenta.Count > 0 Then
                Dim Venta As Double = 0
                Dim Deuda As Double = 0
                Dim Dias As Double = 0

                If dvDeuda.Count > 0 Then
                    Deuda = CDbl(dvDeuda.Item(0).Item("total"))
                End If
                If dvVenta.Count > 0 Then
                    Venta = CDbl(dvVenta.Item(0).Item("total"))
                End If

                If Venta > 0 Then
                    Dias = (Deuda * 365) / Venta
                End If

                tmp.Nuevo(usr.Codigo, "DCALLE")
                tmp.Cadena(0) = rep
                tmp.Numero(0) = Deuda
                tmp.Numero(1) = Venta
                tmp.Numero(2) = Dias

            End If

        Next

        tmp.Grabar()


    End Sub

End Class
