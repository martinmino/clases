Imports System.Data.OracleClient
Imports System.IO

Public Class Parque

    Private cn As OracleConnection

    Private WithEvents da1 As OracleDataAdapter    'machines
    Private WithEvents da2 As OracleDataAdapter    'ymacitm
    Private da3 As OracleDataAdapter    'bomd
    Private da4 As OracleDataAdapter 'macitn

    Private dt1 As DataTable    'machines
    Private dt2 As DataTable    'ymacitm
    Private dt3 As DataTable    'bomd
    Private dt4 As DataTable    'macitn

    Public Event Vencido(ByVal sender As Object, ByVal e As EventArgs)
    Public Event CambioPolvo(ByVal sender As Object, ByVal e As EventArgs)
    Public Event TarjetaImpresa(ByVal sender As Object, ByVal e As EventArgs)
    Public Event CambioFabricacion(ByVal sender As Object, ByVal e As EventArgs)
    Public Event CambioArticulo(ByVal sender As Object, ByVal e As CambioArticuloEventArgs)
    Public Event CambioVencimiento(ByVal sender As Object, ByVal e As CambioVencimientoEvenArgs)

    'NEW
    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
        Adaptadores()
    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal Serie As String)
        Me.New(cn)
        Abrir(Serie)
    End Sub

    'EVENTOS
    Private Sub da1_RowUpdating(ByVal sender As Object, ByVal e As System.Data.OracleClient.OracleRowUpdatingEventArgs) Handles da1.RowUpdating
        If e.Status = UpdateStatus.ErrorsOccurred Then Exit Sub

        If e.StatementType = StatementType.Update Then
            e.Row.BeginEdit()
            e.Row("upddat_0") = Date.Today
            e.Row("updusr_0") = USER
            e.Row.EndEdit()
        End If
    End Sub
    Private Sub da2_RowUpdated(ByVal sender As Object, ByVal e As System.Data.OracleClient.OracleRowUpdatedEventArgs) Handles da2.RowUpdated
        If e.Status = UpdateStatus.ErrorsOccurred Then Exit Sub

        Dim dr As DataRow
        dr = dt1.Rows(0)

        If e.StatementType = StatementType.Update Then
            dr.BeginEdit()
            dr("upddat_0") = Date.Today
            dr("updusr_0") = USER
            dr.EndEdit()
        End If

    End Sub

    'SUB
    Private Sub Adaptadores()
        Dim Sql As String

        'TABLA MACHINES
        Sql = "SELECT * FROM machines WHERE macnum_0 = :macnum_0"
        da1 = New OracleDataAdapter(Sql, cn)
        da1.SelectCommand.Parameters.Add("macnum_0", OracleType.VarChar)
        da1.InsertCommand = New OracleCommandBuilder(da1).GetInsertCommand
        da1.UpdateCommand = New OracleCommandBuilder(da1).GetUpdateCommand
        da1.DeleteCommand = New OracleCommandBuilder(da1).GetDeleteCommand

        'Sql = "UPDATE machines "
        'Sql &= "SET salfcy_0=:salfcy_0, cur_0=:cur_0, macpdtcod_0=:macpdtcod_0, macqty_0=:macqty_0, macsernum_0=:macsernum_0, macbra_0=:macbra_0, macbracla_0=:macbracla_0, "
        'Sql &= "macdes_0=:macdes_0, macitsdat_0=:macitsdat_0, macitntyp_0=:macitntyp_0, maccutbpc_0=:maccutbpc_0, bpcnum_0=:bpcnum_0, ccnnum_0=:ccnnum_0, "
        'Sql &= "macpurdat_0=:macpurdat_0, macrsl_0=:macrsl_0, macitndat_0=:macitndat_0, macsalpri_0=:macsalpri_0, macbpcpri_0=:macbpcpri_0, macbpccur_0=:macbpccur_0, "
        'Sql &= "macbpcdat_0=:macbpcdat_0, macitnlnd_0=:macitnlnd_0, fcyitn_0=:fcyitn_0, ple_0=:ple_0, macori_0=:macori_0, macoritxt_0=:macoritxt_0, macorivcr_0=:macorivcr_0, "
        'Sql &= "macorivcrl_0=:macorivcrl_0, preori_0=:preori_0, preorivcr_0=:preorivcr_0, preorivcrl_0=:preorivcrl_0, creusr_0=:creusr_0, credat_0=:credat_0, updusr_0=:updusr_0, "
        'Sql &= "upddat_0=:upddat_0, ynrocil_0=:ynrocil_0, yfabdat_0=:yfabdat_0, ymaccob_0=:ymaccob_0, xitn_0=:xitn_0, xbajamotiv_0=:xbajamotiv_0, xbajaobs_0=:xbajaobs_0, "
        'Sql &= "recargador_0 = :recargador_0, patente_0 = :patente_0, tipomanga_0 = :tipomanga_0, lngnomi_0 = :lngnomi_0, lngreal_0 = :lngreal_0, diametro_0 = :diametro_0 "
        'Sql &= "WHERE macnum_0=:macnum_0"
        'da1.UpdateCommand = New OracleCommand(Sql, cn)

        'Sql = "INSERT INTO machines "
        'Sql &= "VALUES(:macnum_0, :salfcy_0, :cur_0, :macpdtcod_0, :macqty_0, :macsernum_0, :macbra_0, :macbracla_0, :macdes_0, :macitsdat_0, :macitntyp_0, :maccutbpc_0, "
        'Sql &= ":bpcnum_0, :ccnnum_0, :macpurdat_0, :macrsl_0, :macitndat_0, :macsalpri_0, :macbpcpri_0, :macbpccur_0, :macbpcdat_0, :macitnlnd_0, :fcyitn_0, :ple_0, "
        'Sql &= ":macori_0, :macoritxt_0, :macorivcr_0, :macorivcrl_0, :preori_0, :preorivcr_0, :preorivcrl_0, :creusr_0, :credat_0, :updusr_0, :upddat_0, :ynrocil_0, "
        'Sql &= ":yfabdat_0, :ymaccob_0, :xitn_0, :xbajamotiv_0, :xbajaobs_0, :recargador_0, :patente_0, :tipomanga_0, :lngnomi_0, :lngreal_0, :diametro_0) "
        'da1.InsertCommand = New OracleCommand(Sql, cn)

        'Sql = "DELETE FROM machines WHERE macnum_0 = :macnum_0"
        'da1.DeleteCommand = New OracleCommand(Sql, cn)

        'With da1
        '    .SelectCommand.Parameters.Add("macnum_0", OracleType.VarChar)

        '    Parametro(.UpdateCommand, "macnum_0", OracleType.VarChar, DataRowVersion.Original)
        '    Parametro(.UpdateCommand, "salfcy_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "cur_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "macpdtcod_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "macqty_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "macsernum_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "macbra_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "macbracla_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "macdes_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "macitsdat_0", OracleType.DateTime, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "macitntyp_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "maccutbpc_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "bpcnum_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "ccnnum_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "macpurdat_0", OracleType.DateTime, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "macrsl_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "macitndat_0", OracleType.DateTime, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "macsalpri_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "macbpcpri_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "macbpccur_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "macbpcdat_0", OracleType.DateTime, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "macitnlnd_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "fcyitn_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "ple_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "macori_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "macoritxt_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "macorivcr_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "macorivcrl_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "preori_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "preorivcr_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "preorivcrl_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "creusr_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "credat_0", OracleType.DateTime, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "updusr_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "upddat_0", OracleType.DateTime, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "ynrocil_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "yfabdat_0", OracleType.DateTime, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "ymaccob_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "xitn_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "xbajamotiv_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "xbajaobs_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "recargador_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "patente_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "tipomanga_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "lngnomi_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "lngreal_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "diametro_0", OracleType.Number, DataRowVersion.Current)

        '    Parametro(.InsertCommand, "macnum_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "salfcy_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "cur_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "macpdtcod_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "macqty_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "macsernum_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "macbra_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "macbracla_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "macdes_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "macitsdat_0", OracleType.DateTime, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "macitntyp_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "maccutbpc_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "bpcnum_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "ccnnum_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "macpurdat_0", OracleType.DateTime, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "macrsl_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "macitndat_0", OracleType.DateTime, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "macsalpri_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "macbpcpri_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "macbpccur_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "macbpcdat_0", OracleType.DateTime, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "macitnlnd_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "fcyitn_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "ple_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "macori_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "macoritxt_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "macorivcr_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "macorivcrl_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "preori_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "preorivcr_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "preorivcrl_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "creusr_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "credat_0", OracleType.DateTime, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "updusr_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "upddat_0", OracleType.DateTime, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "ynrocil_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "yfabdat_0", OracleType.DateTime, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "ymaccob_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "xitn_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "xbajamotiv_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "xbajaobs_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "recargador_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "patente_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "tipomanga_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "lngnomi_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "lngreal_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "diametro_0", OracleType.Number, DataRowVersion.Current)

        '    Parametro(.DeleteCommand, "macnum_0", OracleType.VarChar, DataRowVersion.Original)

        'End With

        'TABLA YMACITM
        Sql = "SELECT ymc.* "
        Sql &= "FROM (machines mac INNER JOIN ymacitm ymc ON (mac.macnum_0 = ymc.macnum_0)) INNER JOIN bomd bod ON (mac.macpdtcod_0 = bod.itmref_0 AND ymc.cpnitmref_0 =  bod.cpnitmref_0) "
        Sql &= "WHERE mac.macnum_0 = :macnum_0 "
        Sql &= "ORDER BY bomseq_0"
        da2 = New OracleDataAdapter(Sql, cn)

        Sql = "UPDATE ymacitm "
        Sql &= "SET cpnitmref_0 = :cpnitmref_0, datprev_0 = :datprev_0, datnext_0 = :datnext_0, bpcnum_0 = :bpcnum_0, cllnum_0 = :cllnum_0 "
        Sql &= "WHERE macnum_0 = :macnum_0 AND cpnitmref_0 = :cpnitmref_0w"
        da2.UpdateCommand = New OracleCommand(Sql, cn)

        Sql = "INSERT INTO ymacitm VALUES(:macnum_0, :cpnitmref_0, :datprev_0, :datnext_0, :bpcnum_0, :cllnum_0)"
        da2.InsertCommand = New OracleCommand(Sql, cn)

        Sql = "DELETE FROM ymacitm WHERE macnum_0 = :macnum_0 AND cpnitmref_0 = :cpnitmref_0"
        da2.DeleteCommand = New OracleCommand(Sql, cn)

        With da2
            .SelectCommand.Parameters.Add("macnum_0", OracleType.VarChar)

            Parametro(.UpdateCommand, "cpnitmref_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "datprev_0", OracleType.DateTime, DataRowVersion.Current)
            Parametro(.UpdateCommand, "datnext_0", OracleType.DateTime, DataRowVersion.Current)
            Parametro(.UpdateCommand, "bpcnum_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "cllnum_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "macnum_0", OracleType.VarChar, DataRowVersion.Original)
            Parametro(.UpdateCommand, "cpnitmref_0w", OracleType.VarChar, DataRowVersion.Original, "cpnitmref_0")

            Parametro(.InsertCommand, "macnum_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "cpnitmref_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "datprev_0", OracleType.DateTime, DataRowVersion.Current)
            Parametro(.InsertCommand, "datnext_0", OracleType.DateTime, DataRowVersion.Current)
            Parametro(.InsertCommand, "bpcnum_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "cllnum_0", OracleType.VarChar, DataRowVersion.Current)

            Parametro(.DeleteCommand, "macnum_0", OracleType.VarChar, DataRowVersion.Original)
            Parametro(.DeleteCommand, "cpnitmref_0", OracleType.VarChar, DataRowVersion.Original)

        End With

        'TABLA BOMD
        Sql = "SELECT bomseq_0, cpnitmref_0, ydayfreq_0 FROM bomd WHERE bomalt_0 = 99 AND itmref_0 = :itmref_0 ORDER BY bomseq_0"
        da3 = New OracleDataAdapter(Sql, cn)
        da3.SelectCommand.Parameters.Add("itmref_0", OracleType.VarChar)

        'TABLA MACITN
        Sql = "SELECT * FROM macitn WHERE macnum_0 = :macnum_0"
        da4 = New OracleDataAdapter(Sql, cn)
        da4.SelectCommand.Parameters.Add("macnum_0", OracleType.VarChar)
        da4.InsertCommand = New OracleCommandBuilder(da4).GetInsertCommand

        Sql = "DELETE FROM macitn "
        Sql &= "WHERE macnum_0 = :macnum_0 AND itntyp_0 = :itntyp_0 AND bpc_0 = :bpc_0 AND itnfcy_0 = :itnfcy_0 AND "
        Sql &= "      ple_0 = :ple_0 AND bpctyp_0 = :bpctyp_0 AND purdat_0 = :purdat_0 AND rsl_0 = :rsl_0 AND "
        Sql &= "      itndat_0 = :itndat_0 AND enditn_0 = :enditn_0 AND secpri_0 = :secpri_0 AND salpri_0 = :salpri_0 AND "
        Sql &= "      lnd_0 = :lnd_0 AND cur_0 = :cur_0 AND seccur_0 = :seccur_0 AND bpcdat_0 = :bpcdat_0 AND "
        Sql &= "      bpcpri_0 = :bpcpri_0 AND bpccur_0 = :bpccur_0 AND ori_0 = :ori_0 AND oritxt_0 = :oritxt_0 AND "
        Sql &= "      orivcr_0 = :orivcr_0 AND orivcrl_0 = :orivcrl_0 AND clsnum_0 = :clsnum_0 AND xbajamotiv_0 = :xbajamotiv_0 AND "
        Sql &= "      xbajaobs_0 = :xbajaobs_0 AND tipomanga_0 = :tipomanga_0 AND lngnomi_0 = :lngnomi_0 AND lngreal_0 = :lngreal_0 AND "
        Sql &= "      diametro_0 = :diametro_0 AND sello_0 = :sello_0 "
        da4.DeleteCommand = New OracleCommand(Sql, cn)

        Parametro(da4.DeleteCommand, "macnum_0", OracleType.VarChar, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "itntyp_0", OracleType.Number, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "bpc_0", OracleType.VarChar, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "itnfcy_0", OracleType.VarChar, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "ple_0", OracleType.VarChar, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "bpctyp_0", OracleType.VarChar, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "purdat_0", OracleType.DateTime, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "rsl_0", OracleType.VarChar, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "itndat_0", OracleType.DateTime, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "enditn_0", OracleType.DateTime, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "secpri_0", OracleType.Number, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "salpri_0", OracleType.Number, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "lnd_0", OracleType.Number, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "cur_0", OracleType.VarChar, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "seccur_0", OracleType.VarChar, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "bpcdat_0", OracleType.DateTime, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "bpcpri_0", OracleType.Number, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "bpccur_0", OracleType.VarChar, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "ori_0", OracleType.Number, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "oritxt_0", OracleType.VarChar, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "orivcr_0", OracleType.VarChar, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "orivcrl_0", OracleType.Number, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "clsnum_0", OracleType.Number, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "xbajamotiv_0", OracleType.Number, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "xbajaobs_0", OracleType.VarChar, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "tipomanga_0", OracleType.Number, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "lngnomi_0", OracleType.Number, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "lngreal_0", OracleType.Number, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "diametro_0", OracleType.Number, DataRowVersion.Original)
        Parametro(da4.DeleteCommand, "sello_0", OracleType.Number, DataRowVersion.Original)

    End Sub
    Private Sub LimpiarTablaImplantacion()
        Dim dr As DataRow

        For Each dr In dt4.Rows
            If dr.RowState <> DataRowState.Deleted Then dr.Delete()
        Next
        da4.Update(dt4)

    End Sub
    Public Sub Grabar()
        'Si es nuevo obtengo nuevo numero de serie
        If dt1.Rows(0).RowState = DataRowState.Added Then
            Dim Art As New Articulo(cn, ArticuloCodigo)
            Dim NuevoSerie As String = NuevoNumeroSerie()

            With dt1.Rows(0)
                .BeginEdit()
                .Item("macnum_0") = NuevoSerie
                If Art.GestionSerie = 3 Then .Item("macsernum_0") = NuevoSerie
                If .Item("ynrocil_0").ToString.Trim = "" Then .Item("ynrocil_0") = NuevoSerie
                .EndEdit()
            End With

            For Each dr As DataRow In dt2.Rows
                dr.BeginEdit()
                dr("macnum_0") = NuevoSerie
                dr.EndEdit()
            Next

        End If

        da2.Update(dt2)
        da1.Update(dt1)
        If dt4 IsNot Nothing Then da4.Update(dt4)

    End Sub
    Public Sub Rechazar(ByVal itn As Intervencion, ByVal Motivo As Integer, ByVal Obs As String, ByVal TipoManga As Integer, ByVal LngNominal As Integer, ByVal LngReal As Integer, ByVal Diametro As Integer, ByVal Sello As Integer)
        Dim dr As DataRow
        Dim cm As OracleCommand
        Dim Sql As String

        LimpiarTablaImplantacion()

        dr = dt4.NewRow

        dr("macnum_0") = Me.Serie
        dr("itntyp_0") = 1
        dr("bpc_0") = ClienteNumero
        dr("itnfcy_0") = SucursalNumero
        dr("ple_0") = " "
        dr("bpctyp_0") = "BPC"
        dr("purdat_0") = #12/31/1599#
        dr("rsl_0") = " "
        dr("itndat_0") = Date.Today
        dr("enditn_0") = Date.Today
        dr("secpri_0") = 0
        dr("salpri_0") = 0
        dr("lnd_0") = 1
        dr("cur_0") = " "
        dr("seccur_0") = " "
        dr("bpcdat_0") = #12/31/1599#
        dr("bpcpri_0") = 0
        dr("bpccur_0") = " "
        dr("ori_0") = 1
        dr("oritxt_0") = "Creación manual"
        dr("orivcr_0") = itn.Numero
        dr("orivcrl_0") = 0
        dr("clsnum_0") = 0
        dr("xbajamotiv_0") = Motivo
        dr("xbajaobs_0") = IIf(Obs = "", " ", Obs)
        dr("tipomanga_0") = TipoManga
        dr("lngnomi_0") = LngNominal
        dr("lngreal_0") = LngReal
        dr("diametro_0") = Diametro
        dr("sello_0") = Sello
        dt4.Rows.Add(dr)

        dr = dt1.Rows(0)
        dr.BeginEdit()
        dr("MACITNTYP_0") = 5
        dr("MACCUTBPC_0") = " "
        dr("MACITNDAT_0") = Date.Today
        dr("MACORI_0") = 6
        dr("MACORITXT_0") = "Modificación manual de la implantación"
        dr("MACORIVCR_0") = dr("macnum_0").ToString
        dr("UPDUSR_0") = USER
        dr("UPDDAT_0") = Date.Today
        dr.EndEdit()

        'Si el equipo figura en algun puesto de abonados, dejo el puesto vacio
        Sql = "UPDATE xpuestos SET macnum_0 = ' ', estado_0 = '00' WHERE macnum_0 = :macnum_0"
        cm = New OracleCommand(Sql, cn)
        cm.Parameters.Add("macnum_0", OracleType.VarChar).Value = Me.Serie
        cm.ExecuteNonQuery()

        cm.Dispose() : cm = Nothing

    End Sub
    Public Sub QuitarRechazo()
        Dim dr As DataRow = dt1.Rows(0)

        dr.BeginEdit()
        dr("maccutbpc_0") = dr("bpcnum_0").ToString
        dr("macitntyp_0") = 1
        dr.EndEdit()

        LimpiarTablaImplantacion()

    End Sub
    Public Sub Deshacer()
        dt1.RejectChanges()
        dt2.RejectChanges()
    End Sub
    Public Sub Procesar(ByVal Fecha As Date, Optional ByVal Cancel As Boolean = False)
        Dim dr As DataRow = dt1.Rows(0)

        If Not Cancel Then
            dr.BeginEdit()
            dr("xitn_0") = " "
            dr.EndEdit()
        End If

        'Actualizar fecha de recarga
        If dt2.Rows.Count >= 1 Then
            With dt2.Rows(0)
                .BeginEdit()
                .Item("datprev_0") = Fecha.AddDays(FrecuenciaRecarga)
                .Item("datnext_0") = Fecha.AddDays(FrecuenciaRecarga)
                .EndEdit()
            End With
        End If

        'Actualizar fecha de PH
        If dt2.Rows.Count = 2 And dt3.Rows.Count = 2 Then
            With dt2.Rows(1)
                If CDate(.Item("datnext_0")) < Fecha.AddYears(1) Then
                    .BeginEdit()
                    .Item("datprev_0") = Fecha.AddDays(FrecuenciaPH)
                    .Item("datnext_0") = Fecha.AddDays(FrecuenciaPH)
                    .EndEdit()
                End If
            End With
        End If

    End Sub
    Public Sub ProcesarExtintor()
        Dim dr As DataRow = dt1.Rows(0)
        Dim fecha As Date = Date.Today

        dr.BeginEdit()
        dr("xitn_0") = " "
        dr.EndEdit()

        'Actualizar fecha de recarga
        If dt2.Rows.Count >= 1 Then
            With dt2.Rows(0)
                .BeginEdit()
                .Item("datprev_0") = New Date(fecha.Year + 1, fecha.Month, 1)
                .Item("datnext_0") = New Date(fecha.Year + 1, fecha.Month, 1)
                .EndEdit()
            End With
        End If

        'Actualizar fecha de PH
        If dt2.Rows.Count = 2 And dt3.Rows.Count = 2 Then
            With dt2.Rows(1)
                If CDate(.Item("datnext_0")) < fecha.AddYears(1) Then
                    .BeginEdit()
                    .Item("datprev_0") = New Date(fecha.Year + 1, fecha.Month, 1)
                    .Item("datnext_0") = fecha.AddDays(FrecuenciaPH)
                    .EndEdit()
                End If
            End With
        End If

        '---------------------------------------------------------------------------------
        ' Si el equipo es de polvo y el año de fabricacion es igual a (AñoActual-10)
        ' Se dispara evento CambioPolvo
        '---------------------------------------------------------------------------------
        If dr("macpdtcod_0").ToString.Substring(2, 1) = "2" Then
            If FabricacionCorto = fecha.Year - 10 Then RaiseEvent CambioPolvo(Me, New EventArgs)
        End If

    End Sub
    Public Sub ImprimirRechazo(ByVal itn As Intervencion, ByVal Pallet As Long, ByVal Path As String, ByVal Puerto As String)
        Dim dr As DataRow = dt1.Rows(0)
        Dim Etiqueta As String
        Dim Archivo As String = String.Format("{0}\{1}.txt", Path, Environment.MachineName)
        Dim st As Stream
        Dim sr As StreamReader
        Dim sw As StreamWriter

        'Cargo modelo de etiqueta segun tipo de parque
        st = File.Open(Path & "\rechazo.txt", FileMode.Open, FileAccess.Read, FileShare.Read)

        sr = New StreamReader(st)
        Etiqueta = sr.ReadToEnd
        sr.Close()
        st.Close()

        With Cliente
            Etiqueta = Etiqueta.Replace("{PALLET}", Pallet.ToString("N0"))
            Etiqueta = Etiqueta.Replace("{ITN}", itn.Numero)
            Etiqueta = Etiqueta.Replace("{OT}", IIf(itn.OTR = 0, "", "(" & itn.OTR.ToString & ")").ToString)
            Etiqueta = Etiqueta.Replace("{NRO}", Cilindro)
            Etiqueta = Etiqueta.Replace("{TEXTO}", RechazoMotivoTexto)
        End With

        'Grabo archivo
        st = File.Open(Archivo, FileMode.Create, FileAccess.Write, FileShare.None)
        sw = New StreamWriter(st)
        sw.WriteLine(Etiqueta)
        sw.Close()
        st.Close()

        'Copio archivo a puerto paralelo
        If USER <> "MMIN" Then File.Copy(Archivo, Puerto)

        RaiseEvent TarjetaImpresa(Me, New EventArgs)

    End Sub
    Public Sub ImprimirEtiqueta(ByVal Intervencion As String, ByVal Path As String, ByVal Puerto As String)
        Dim dr As DataRow = dt1.Rows(0)

        Dim Etiqueta As String
        Dim Archivo As String = String.Format("{0}\{1}.txt", Path, Environment.MachineName)
        Dim st As Stream
        Dim sr As StreamReader
        Dim sw As StreamWriter

        'Cargo modelo de etiqueta segun tipo de parque

        If EsManguera Then
            st = File.Open(Path & "\etiqueta_manguera.txt", FileMode.Open, FileAccess.Read, FileShare.Read)

        Else
            st = File.Open(Path & "\etiqueta_extintor.txt", FileMode.Open, FileAccess.Read, FileShare.Read)

        End If

        sr = New StreamReader(st)
        Etiqueta = sr.ReadToEnd
        sr.Close()
        st.Close()

        With Cliente
            Etiqueta = Etiqueta.Replace("%codigo_cliente%", .Codigo)
            Etiqueta = Etiqueta.Replace("%nombre_cliente%", .Nombre)
        End With

        If EsManguera Then
            Etiqueta = Etiqueta.Replace("%vto_ph%", VtoCarga.ToString("MM / yy"))
        Else
            Etiqueta = Etiqueta.Replace("%vto_ph%", VtoPH.ToString("MM / yy"))
        End If

        Etiqueta = Etiqueta.Replace("%codigo_articulo%", Articulo.FamiliaDescripcion(2) & " " & Articulo.FamiliaDescripcion(1)) 'dr("macpdtcod_0").ToString)
        Etiqueta = Etiqueta.Replace("%descripcion_articulo%", "") 'dr("macdes_0").ToString.Replace("""", "\"""))
        Etiqueta = Etiqueta.Replace("%intervencion%", Intervencion)
        Etiqueta = Etiqueta.Replace("%serie%", Serie)
        Etiqueta = Etiqueta.Replace("%cilindro%", Cilindro)
        Etiqueta = Etiqueta.Replace("%fecha_actual%", Date.Now.ToString("dd/MM/yy HH:mm"))
        Etiqueta = Etiqueta.Replace("%fabricacion%", FabricacionCorto.ToString)

        'Grabo archivo
        st = File.Open(Archivo, FileMode.Create, FileAccess.Write, FileShare.None)
        sw = New StreamWriter(st)
        sw.WriteLine(Etiqueta)
        sw.Close()
        st.Close()

        'Copio archivo a puerto paralelo
        File.Copy(Archivo, Puerto)

        RaiseEvent TarjetaImpresa(Me, New EventArgs)

    End Sub
    Public Sub Nuevo(ByVal Cliente As String, ByVal Sucursal As String)
        Dim dr As DataRow

        If dt1 Is Nothing Then
            Dim da As New OracleDataAdapter("SELECT * FROM ymacitm WHERE macnum_0 = ',,-.,'", cn)

            dt1 = New DataTable
            dt2 = New DataTable
            dt3 = New DataTable
            dt4 = New DataTable

            da1.FillSchema(dt1, SchemaType.Mapped)
            da.FillSchema(dt2, SchemaType.Mapped)
            da3.FillSchema(dt3, SchemaType.Mapped)

            da.Dispose()

        Else
            dt1.Clear()
            dt2.Clear()
            dt3.Clear()
            dt4.Clear()

        End If

        dr = dt1.NewRow
        dr("macnum_0") = " "
        dr("salfcy_0") = " "
        dr("cur_0") = " "
        dr("macpdtcod_0") = " "
        dr("macqty_0") = 1
        dr("macsernum_0") = " "
        dr("macbra_0") = "01"
        dr("macbracla_0") = "GEORGIA"
        dr("macdes_0") = " "
        dr("macitsdat_0") = #12/31/1599#
        dr("macitntyp_0") = 1
        dr("maccutbpc_0") = Cliente
        dr("bpcnum_0") = Cliente
        dr("ccnnum_0") = " "
        dr("macpurdat_0") = #12/31/1599#
        dr("macrsl_0") = " "
        dr("macitndat_0") = #12/31/1599#
        dr("macsalpri_0") = 0
        dr("macbpcpri_0") = 0
        dr("macbpccur_0") = " "
        dr("macbpcdat_0") = #12/31/1599#
        dr("macitnlnd_0") = 1
        dr("fcyitn_0") = Sucursal
        dr("ple_0") = " "
        dr("macori_0") = 1
        dr("macoritxt_0") = "Creacion manual"
        dr("macorivcr_0") = " "
        dr("macorivcrl_0") = 0
        dr("preori_0") = 1
        dr("preorivcr_0") = " "
        dr("preorivcrl_0") = 0
        dr("creusr_0") = USER
        dr("credat_0") = Date.Today
        dr("updusr_0") = " "
        dr("upddat_0") = #12/31/1599#
        dr("ynrocil_0") = " "
        dr("yfabdat_0") = #12/31/1599#
        dr("ymaccob_0") = 1
        dr("xitn_0") = " "
        dr("xbajamotiv_0") = 0
        dr("xbajaobs_0") = " "
        dr("recargador_0") = 0
        dr("patente_0") = " "
        dr("tipomanga_0") = 0
        dr("lngnomi_0") = 0
        dr("lngreal_0") = 0
        dr("diametro_0") = 0
        dr("nromanga_0") = 0
        dt1.Rows.Add(dr)

    End Sub
    Private Sub SetArticulo(ByVal Codigo As String)

        Dim dr As DataRow = dt1.Rows(0)
        Dim i, j As Integer
        Dim Arti = New Articulo(cn, Codigo)

        'Cambio el codigo de articulo
        If Codigo = dr("macpdtcod_0").ToString Then Exit Sub

        dr.BeginEdit()
        dr("macpdtcod_0") = Codigo
        dr("macdes_0") = Arti.Descripcion
        If Arti.GestionSerie = 3 Then
            dr("macsernum_0") = dr("macnum_0")
        Else
            dr("macsernum_0") = " "
        End If
        dr.EndEdit()

        'Busco los vencimientos del articulo
        da3.SelectCommand.Parameters("itmref_0").Value = Codigo
        dt3.Clear()
        da3.Fill(dt3)

        'Recorro los vencimientos encontrados
        For i = 0 To dt3.Rows.Count - 1

            If dt2.Rows.Count - 1 < i Then
                'Agrego registro de vencimiento de servicio
                dr = dt2.NewRow
                dr("macnum_0") = Serie
                dr("cpnitmref_0") = dt3.Rows(i).Item("cpnitmref_0").ToString
                dr("datprev_0") = Date.Today.AddDays(CInt(dt3.Rows(i).Item("ydayfreq_0")))
                dr("datnext_0") = Date.Today.AddDays(CInt(dt3.Rows(i).Item("ydayfreq_0")))
                dr("bpcnum_0") = ClienteNumero
                dr("cllnum_0") = " "
                dt2.Rows.Add(dr)

                RaiseEvent CambioVencimiento(Me, New CambioVencimientoEvenArgs(CDate(dr("datnext_0")), i))

            Else
                With dt2.Rows(i)
                    If .RowState = DataRowState.Deleted Then .RejectChanges()

                    If .Item("cpnitmref_0").ToString <> dt3.Rows(i).Item("cpnitmref_0").ToString Then
                        'Modifico el codigo de servicio
                        .BeginEdit()
                        .Item("cpnitmref_0") = dt3.Rows(i).Item("cpnitmref_0").ToString
                        .EndEdit()
                    End If
                End With

            End If

        Next

        'Elimino los registros que pudieran existir
        For j = i To dt2.Rows.Count - 1
            dr = dt2.Rows(i)
            dr.Delete()
        Next

        RaiseEvent CambioArticulo(Me, New CambioArticuloEventArgs(dt2.Rows.Count = 2))

    End Sub
    Public Sub Borrar()
        Dim dr As DataRow

        For Each dr In dt1.Rows
            dr.BeginEdit()
            dr.Delete()
            dr.EndEdit()
        Next
        For Each dr In dt2.Rows
            dr.BeginEdit()
            dr.Delete()
            dr.EndEdit()
        Next

        Grabar()

    End Sub
    Public Sub setTipoExtintor(ByVal Agente As String, ByVal Capacidad As String)
        Dim da As OracleDataAdapter
        Dim sql As String
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim art As String = "112016"

        sql = "SELECT itmref_0 " _
            & "FROM itmmaster " _
            & "WHERE tsicod_2 = :agente AND " _
            & "      tsicod_1 = :capacidad " _
            & "ORDER BY itmref_0"

        da = New OracleDataAdapter(sql, cn)
        da.SelectCommand.Parameters.Add("agente", OracleType.VarChar).Value = Agente
        da.SelectCommand.Parameters.Add("capacidad", OracleType.VarChar).Value = Capacidad

        da.Fill(dt)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            art = dr(0).ToString
        End If

        SetArticulo(art)

        da.Dispose()
        dt.Dispose()

    End Sub

    'FUNCTION
    Public Function Abrir(ByVal Serie As String) As Boolean
        Dim dr As DataRow

        da1.SelectCommand.Parameters("macnum_0").Value = Serie
        da2.SelectCommand.Parameters("macnum_0").Value = Serie
        da4.SelectCommand.Parameters("macnum_0").Value = Serie

        If dt1 Is Nothing Then
            dt1 = New DataTable
            dt2 = New DataTable
            dt3 = New DataTable
            dt4 = New DataTable

        Else
            dt1.Clear()
            dt2.Clear()
            dt3.Clear()
            dt4.Clear()

        End If


        Try
            da1.Fill(dt1)
            da2.Fill(dt2)
            da4.Fill(dt4)

            If dt1.Rows.Count = 1 Then
                dr = dt1.Rows(0)

                da3.SelectCommand.Parameters("itmref_0").Value = dr("macpdtcod_0").ToString
                da3.Fill(dt3)

                Return True

            Else
                Return False

            End If

        Catch ex As Exception

            Return False

        End Try

    End Function
    Private Function NuevoNumeroSerie() As String
        Dim Serie As Integer
        Dim Sql = "SELECT valeur_0 FROM avalnum where codnum_0 = 'MAC'"
        Dim cm As New OracleCommand(Sql, cn)
        Dim dr As OracleDataReader
        Dim s As String = ""

        dr = cm.ExecuteReader(CommandBehavior.SingleResult)
        dr.Read()
        Serie = CInt(dr(0))
        dr.Close()

        Sql = String.Format("UPDATE avalnum SET valeur_0 = {0} WHERE codnum_0 = 'MAC'", Serie + 1)
        cm.CommandText = Sql
        cm.ExecuteNonQuery()

        cm.Dispose()

        s = String.Format("00000000{0}", Serie)

        Return s.Substring(s.Length - 8)

    End Function
    Public Function ExisteCilindro(ByVal Cilindro As String) As Boolean
        Dim da As New OracleDataAdapter("SELECT * FROM machines WHERE macnum_0 <> :macnum_0 AND bpcnum_0 = :bpcnum_0 AND ynrocil_0 = :ynrocil_0", cn)
        Dim dt As New DataTable

        da.SelectCommand.Parameters.Add("macnum_0", OracleType.VarChar).Value = Serie
        da.SelectCommand.Parameters.Add("bpcnum_0", OracleType.VarChar).Value = ClienteNumero
        da.SelectCommand.Parameters.Add("ynrocil_0", OracleType.VarChar).Value = Cilindro

        Try
            'Recupero los cilindros
            da.Fill(dt)
            ExisteCilindro = (dt.Rows.Count > 0)

        Catch ex As Exception
            'MessageBox.Show(ex.Message, "Parque.ExisteCilindro()", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

        da.Dispose()
        dt.Dispose()

    End Function
    Public Function SimularVidaUtil(ByVal Fabricacion As Integer) As Boolean
        'Recibe un año de fabricacion y devuelve si esta vencido para ese año
        Fabricacion += TiempoVidaUtil
        Return (Date.Today.Year >= Fabricacion)
    End Function
    Public Function ObtenerParquePrestamo(ByVal Nro As String, ByVal Articulo As String) As String
        Dim da As New OracleDataAdapter("SELECT * FROM machines WHERE ynrocil_0 = :ynrocil_0 AND macpdtcod_0 = :macpdtcod_0", cn)
        Dim dt As New DataTable

        da.SelectCommand.Parameters.Add("ynrocil_0", OracleType.VarChar).Value = Nro
        da.SelectCommand.Parameters.Add("macpdtcod_0", OracleType.VarChar).Value = Articulo
        da.Fill(dt)

        If dt.Rows.Count = 1 Then
            Nro = dt.Rows(0).Item("macnum_0").ToString
        Else
            Nro = ""
        End If

        da.Dispose()
        dt.Dispose()

        Return Nro

    End Function
    Shared Function ObtenerParque(ByVal cn As OracleConnection, ByVal Articulo As String, ByVal Cliente As String, ByVal Sucursal As String, ByRef Series() As String) As Boolean
        'Devuelve los numeros de serie de todos los parques del cliente del tipo = articulo
        Dim da As New OracleDataAdapter("SELECT macnum_0 FROM machines WHERE bpcnum_0 = :bpcnum_0 AND fcyitn_0 = :fcyitn_0 AND macpdtcod_0 = :macpdtcod_0 ORDER BY macnum_0", cn)
        Dim dt As New DataTable
        Dim i As Integer = -1

        With da.SelectCommand.Parameters
            .Add("bpcnum_0", OracleType.VarChar).Value = Cliente
            .Add("fcyitn_0", OracleType.VarChar).Value = Sucursal
            .Add("macpdtcod_0", OracleType.VarChar).Value = Articulo
        End With

        da.Fill(dt)

        If dt.Rows.Count > 0 Then
            ReDim Series(dt.Rows.Count - 1)

            For i = 0 To dt.Rows.Count - 1
                Series(i) = dt.Rows(i).Item("macnum_0").ToString
            Next

        End If

        dt.Dispose()
        da.Dispose()

        Return (i > -1)

    End Function
    Shared Function ObtenerParque(ByVal cn As OracleConnection, ByVal Cliente As String, ByVal Sucursal As String, ByVal Desde As Date, ByVal Hasta As Date) As DataTable
        'Devuelve los numeros de serie de todos los parques del cliente del tipo = articulo
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim Sql As String

        Sql = "SELECT mac.*, ymc.* "
        Sql &= "FROM (machines mac INNER JOIN ymacitm ymc ON (mac.macnum_0 = ymc.macnum_0)) INNER JOIN bomd bmd ON (bmd.itmref_0 = macpdtcod_0 AND bmd.cpnitmref_0 = ymc.cpnitmref_0) "
        Sql &= "WHERE bomalt_0 = 99 AND "
        Sql &= "	  bomseq_0 = 10 AND "
        Sql &= "	  cpnitmref_0 LIKE '45%' AND "
        Sql &= "	  mac.bpcnum_0 = :bpcnum AND "
        Sql &= "	  fcyitn_0 = :fcyitn and "
        Sql &= "	  macitntyp_0 <> 5 and "
        Sql &= "	  datnext_0 >= to_date(:desde, 'dd/mm/yyyy') AND datnext_0 < to_date(:hasta, 'dd/mm/yyyy')"

        da = New OracleDataAdapter(Sql, cn)

        With da.SelectCommand.Parameters
            .Add("bpcnum", OracleType.VarChar).Value = Cliente
            .Add("fcyitn", OracleType.VarChar).Value = Sucursal
            .Add("desde", OracleType.VarChar).Value = Desde.ToString("dd/MM/yyyy")
            .Add("hasta", OracleType.VarChar).Value = Hasta.ToString("dd/MM/yyyy")
        End With
        da.Fill(dt)
        da.Dispose()

        Return dt

    End Function
    Shared Function ExisteExtructuraComercial(ByVal cn As OracleConnection, ByVal Articulo As String) As Boolean
        'Devuelve los numeros de serie de todos los parques del cliente del tipo = articulo
        Dim da As New OracleDataAdapter("SELECT * FROM bomd WHERE itmref_0 = :itmref_0", cn)
        Dim dt As New DataTable

        da.SelectCommand.Parameters.Add("itmref_0", OracleType.VarChar).Value = Articulo
        da.Fill(dt)
        da.Dispose()

        Return (dt.Rows.Count > 0)

    End Function
    Shared Function ObtenerVtos(ByVal cn As OracleConnection, ByVal Desde As Date, ByVal Hasta As Date, ByVal Cliente As String, ByVal Sucursal As String) As DataTable
        'Devuelve la cantidad de vencimientos dentro del rango de fecha
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim sql As String

        sql = "SELECT datnext_0 "
        sql &= "FROM (machines mac INNER JOIN ymacitm ymc ON (mac.macnum_0 = ymc.macnum_0)) INNER JOIN bomd bmd ON (bmd.itmref_0 = macpdtcod_0 AND bmd.cpnitmref_0 = ymc.cpnitmref_0) "
        sql &= "WHERE bomalt_0 = 99 AND bomseq_0 = 10 AND (cpnitmref_0 LIKE '45%' or mac.macpdtcod_0 = '999003') AND macitntyp_0 <> 5 AND mac.bpcnum_0 = :bpcnum_0 AND fcyitn_0 = :fcyitn_0 AND datnext_0 >= to_date(:desde, 'dd/mm/yyyy') AND datnext_0 < to_date(:hasta, 'dd/mm/yyyy') "
        sql &= "ORDER BY datnext_0"

        da = New OracleDataAdapter(sql, cn)
        da.SelectCommand.Parameters.Add("bpcnum_0", OracleType.VarChar).Value = Cliente
        da.SelectCommand.Parameters.Add("fcyitn_0", OracleType.VarChar).Value = Sucursal
        da.SelectCommand.Parameters.Add("desde", OracleType.VarChar).Value = Desde.ToString("dd/MM/yyyy")
        da.SelectCommand.Parameters.Add("hasta", OracleType.VarChar).Value = Hasta.ToString("dd/MM/yyyy")

        da.Fill(dt)
        da.Dispose()

        Return dt

    End Function
    Shared Sub LimpiarParqueGrupo(ByVal cn As OracleConnection, ByVal cliente As String, ByVal sucursal As String, ByVal articulo As String)
        Dim itm As New Articulo(cn, articulo)
        If itm.Grupo = "" Then Exit Sub 'Salgo si el articulo no tiene grupo

        Dim mac As New Parque(cn)
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim Sql As String
        Dim i As Integer

        Sql = "SELECT macnum_0 "
        Sql &= "FROM machines mac INNER JOIN itmmaster itm ON (mac.macpdtcod_0 = itm.itmref_0) "
        Sql &= "WHERE bpcnum_0 = :bpcnum_0 AND fcyitn_0 = :fcyitn_0 AND xgrp_0 = :xgrp_0 AND macpdtcod_0 <> :macpdtcod_0"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("bpcnum_0", OracleType.VarChar).Value = cliente
        da.SelectCommand.Parameters.Add("fcyitn_0", OracleType.VarChar).Value = sucursal
        da.SelectCommand.Parameters.Add("xgrp_0", OracleType.VarChar).Value = itm.Grupo
        da.SelectCommand.Parameters.Add("macpdtcod_0", OracleType.VarChar).Value = articulo

        da.Fill(dt)
        da.Dispose()

        i = dt.Rows.Count

        For Each dr As DataRow In dt.Rows
            If mac.Abrir(dr(0).ToString) Then mac.Borrar()
        Next

        dt.Dispose() : dt = Nothing
        itm.Dispose() : itm = Nothing

    End Sub
    Shared Sub EliminarPorCodigo(ByVal cn As OracleConnection, ByVal cliente As String, ByVal sucursal As String, ByVal articulo As String)
        Dim mac As New Parque(cn)
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim Sql As String

        Sql = "SELECT macnum_0 "
        Sql &= "FROM machines "
        Sql &= "WHERE bpcnum_0 = :bpcnum_0 AND fcyitn_0 = :fcyitn_0 AND macpdtcod_0 = :macpdtcod_0"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("bpcnum_0", OracleType.VarChar).Value = cliente
        da.SelectCommand.Parameters.Add("fcyitn_0", OracleType.VarChar).Value = sucursal
        da.SelectCommand.Parameters.Add("macpdtcod_0", OracleType.VarChar).Value = articulo

        da.Fill(dt)
        da.Dispose()

        For Each dr As DataRow In dt.Rows
            If mac.Abrir(dr(0).ToString) Then mac.Borrar()
        Next

        dt.Dispose() : dt = Nothing

    End Sub
    Shared Sub EliminarPorContador(ByVal cn As OracleConnection, ByVal cliente As String, ByVal sucursal As String, ByVal Contador As String)
        Dim mac As New Parque(cn)
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim Sql As String

        Sql = "SELECT macnum_0 "
        Sql &= "FROM machines "
        Sql &= "WHERE bpcnum_0 = :bpcnum_0 AND fcyitn_0 = :fcyitn_0 AND macnum_0 = :macnum_0"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("bpcnum_0", OracleType.VarChar).Value = cliente
        da.SelectCommand.Parameters.Add("fcyitn_0", OracleType.VarChar).Value = sucursal
        da.SelectCommand.Parameters.Add("macnum_0", OracleType.VarChar).Value = Contador

        da.Fill(dt)
        da.Dispose()

        For Each dr As DataRow In dt.Rows
            If mac.Abrir(dr(0).ToString) Then mac.Borrar()
        Next

        dt.Dispose() : dt = Nothing
    End Sub
    Shared Sub MoverMarcaIntervencion(ByVal cn As OracleConnection, ByVal IntervencionVieja As String, ByVal IntervencionNueva As String)
        Dim Sql As String
        Dim cm As OracleCommand

        Sql = "UPDATE machines SET xitn_0 = :nueva WHERE xitn_0 = :vieja"
        cm = New OracleCommand(Sql, cn)
        cm.Parameters.Add("nueva", OracleType.VarChar).Value = IntervencionNueva
        cm.Parameters.Add("vieja", OracleType.VarChar).Value = IntervencionVieja

        Try
            If cn.State = ConnectionState.Closed Then cn.Open()
            cm.ExecuteNonQuery()

        Catch ex As Exception
        End Try
        

    End Sub
    'PROPERTY
    Public Property Cantidad() As Integer
        Get
            If dt1.Rows.Count = 0 Then Exit Property

            Dim dr As DataRow = dt1.Rows(0)

            Return CInt(dr("macqty_0"))
        End Get
        Set(ByVal value As Integer)
            If dt1.Rows.Count > 0 Then
                Dim dr As DataRow = dt1.Rows(0)
                dr.BeginEdit()
                dr("macqty_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property VtoCarga() As Date
        Get
            If dt2.Rows.Count >= 1 Then
                Return CDate(dt2.Rows(0).Item("datnext_0"))
            Else
                Return #12/31/1599#
            End If
        End Get
        Set(ByVal value As Date)
            If dt2.Rows.Count >= 1 Then
                With dt2.Rows(0)
                    .BeginEdit()
                    .Item("datnext_0") = value
                    .EndEdit()
                End With
            End If
        End Set
    End Property
    Public Property UltimaCarga() As Date
        Get
            If dt2.Rows.Count >= 1 Then
                Return CDate(dt2.Rows(0).Item("prevnext_0"))
            Else
                Return #12/31/1599#
            End If
        End Get
        Set(ByVal value As Date)
            If dt2.Rows.Count >= 1 Then
                With dt2.Rows(0)
                    .BeginEdit()
                    .Item("prevnext_0") = value
                    .EndEdit()
                End With
            End If
        End Set
    End Property
    Public Property VtoPH() As Date
        Get
            If dt2.Rows.Count > 1 Then
                Return CDate(dt2.Rows(1).Item("datnext_0"))
            Else
                Return #12/31/1599#
            End If

        End Get
        Set(ByVal value As Date)
            If dt2.Rows.Count > 1 Then
                With dt2.Rows(1)
                    .BeginEdit()
                    .Item("datnext_0") = value
                    .EndEdit()
                End With
            End If
        End Set
    End Property
    Public Property FabricacionCorto() As Integer
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return CDate(dr("yfabdat_0")).Year
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow = dt1.Rows(0)

            If value < 1599 Then value = 1599
            Dim d As New Date(value, 1, 15)
            dr.BeginEdit()
            dr("yfabdat_0") = d
            dr.EndEdit()

            RaiseEvent CambioFabricacion(Me, New EventArgs)
        End Set
    End Property
    Public Property FabricacionLargo() As Date
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return CDate(dr("yfabdat_0"))
        End Get
        Set(ByVal value As Date)
            Me.FabricacionCorto = value.Year
        End Set
    End Property
    Public ReadOnly Property Cliente() As Cliente
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Dim bpc As New Cliente(cn)
            bpc.Abrir(dr("bpcnum_0").ToString)
            Return bpc
        End Get
    End Property
    Public Property ClienteNumero() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("bpcnum_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt1.Rows(0)

            dr.BeginEdit()
            dr("bpcnum_0") = value
            dr("maccutbpc_0") = value
            dr.EndEdit()

            For Each dr In dt2.Rows
                dr.BeginEdit()
                dr("bpcnum_0") = value
                dr.EndEdit()
            Next
        End Set
    End Property
    Public Property SucursalNumero() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("fcyitn_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt1.Rows(0)

            If dr("fcyitn_0").ToString <> value Then
                dr.BeginEdit()
                dr("fcyitn_0") = value
                dr("ple_0") = " "
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property Cilindro() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)

            Return dr("ynrocil_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt1.Rows(0)

            dr.BeginEdit()
            dr("ynrocil_0") = EsNumerico(value)
            dr.EndEdit()
        End Set
    End Property
    Public Property Observaciones() As String
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)

            Return dr("ple_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dt1.Rows(0)

            dr.BeginEdit()
            dr("ple_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property UltimoRecargador() As Integer
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)

            Return CInt(dr("recargador_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dt1.Rows(0)

            dr.BeginEdit()
            dr("recargador_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property Articulo() As Articulo
        Get
            Return New Articulo(cn, ArticuloCodigo)
        End Get

    End Property
    Public Property ArticuloCodigo() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("macpdtcod_0").ToString
        End Get
        Set(ByVal value As String)
            SetArticulo(value)
        End Set
    End Property
    Public ReadOnly Property ArticuloDescripcion() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("macdes_0").ToString
        End Get
    End Property
    Public ReadOnly Property Serie() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("macnum_0").ToString
        End Get
    End Property
    Public ReadOnly Property TiempoVidaUtil() As Integer
        Get
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)

                If dr("macpdtcod_0").ToString.Substring(2, 1) = "3" Then
                    Return 30

                Else
                    Return 20

                End If

            Else
                Return 0

            End If

        End Get
    End Property
    Public ReadOnly Property FinVidaUtil() As Integer
        Get
            Dim aa As Integer
            Dim dr As DataRow = dt1.Rows(0)

            aa = TiempoVidaUtil
            Return CDate(dr("yfabdat_0")).Year + aa

        End Get
    End Property
    Public ReadOnly Property SuperaVidaUtil() As Boolean
        Get
            Return (Date.Today.Year >= Me.FinVidaUtil)
        End Get
    End Property
    Public ReadOnly Property FrecuenciaRecarga() As Integer
        Get
            Return CInt(dt3.Rows(0).Item("ydayfreq_0"))
        End Get
    End Property
    Public ReadOnly Property FrecuenciaPH() As Integer
        Get
            Return CInt(dt3.Rows(1).Item("ydayfreq_0"))
        End Get
    End Property
    Public ReadOnly Property Servicio() As Articulo
        Get
            Return New Articulo(cn, dt2.Rows(0).Item("cpnitmref_0").ToString)
        End Get
    End Property
    Public ReadOnly Property EsManguera() As Boolean
        Get
            Return dt2.Rows(0).Item("cpnitmref_0").ToString.StartsWith("50500")
        End Get
    End Property

    Public ReadOnly Property TienePh() As Boolean
        Get
            Return dt2.Rows.Count = 2
        End Get
    End Property
    Public ReadOnly Property LlevaTarjeta() As Boolean
        Get
            Dim Sql As String = "SELECT * FROM itmsales WHERE (yflgiram_0 = 2 OR yflgsat_0 = 2) AND itmref_0 = :itmref_0"
            Dim dt As New DataTable
            Dim da As New OracleDataAdapter(Sql, cn)
            da.SelectCommand.Parameters.Add("itmref_0", OracleType.VarChar).Value = ArticuloCodigo

            da.Fill(dt)
            da.Dispose()

            Return (dt.Rows.Count > 0)

        End Get
    End Property
    Public Property Patente() As String
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)

            Return dr("patente_0").ToString.Trim
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dt1.Rows(0)

            dr.BeginEdit()
            dr("patente_0") = value.ToUpper
            dr.EndEdit()
        End Set
    End Property
    Public Property MarcaIntervencion() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)

            Return dr("xitn_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt1.Rows(0)

            dr.BeginEdit()
            dr("xitn_0") = IIf(value = "", " ", value)
            dr.EndEdit()
        End Set
    End Property

    Public ReadOnly Property Rechazado() As Boolean
        Get
            Dim dr As DataRow = dt1.Rows(0)

            Return CInt(dr("macitntyp_0")) = 5
        End Get
    End Property
    Public ReadOnly Property RechazoMotivo() As Integer
        Get
            Dim i As Integer = 0

            If dt4.Rows.Count > 0 Then
                Dim dr As DataRow = dt4.Rows(0)
                i = CInt(dr("xbajamotiv_0"))
            End If

            Return i
        End Get
    End Property
    Public ReadOnly Property RechazoMotivoTexto() As String
        Get
            Dim s As String = ""

            If dt4.Rows.Count > 0 Then
                Dim dr As DataRow = dt4.Rows(0)
                Dim m As New MenuLocal(cn, 2410, False)
                s = m.Descripcion(CInt(dr("xbajamotiv_0")))
            End If

            Return s
        End Get
    End Property
    Public ReadOnly Property RechazoObservacion() As String
        Get
            Dim s As String = ""

            If dt4.Rows.Count > 0 Then
                Dim dr As DataRow = dt4.Rows(0)
                s = dr("xbajaobs_0").ToString
            End If

            Return s
        End Get
    End Property
    Shared Function EsNumerico(ByVal value As String) As String

        For i As Integer = 0 To value.Length - 1
            If Not IsNumeric(value.Substring(i)) Then
                Return value
            End If
        Next
        Return (CLng(value)).ToString

    End Function
End Class

Public Class CambioArticuloEventArgs
    Inherits System.EventArgs

    Public TienePH As Boolean

    Public Sub New(ByVal PH As Boolean)
        Me.TienePH = PH
    End Sub

End Class
Public Class CambioVencimientoEvenArgs
    Inherits System.EventArgs

    Public Fecha As Date
    Public TipoVencimiento As Integer

    Public Sub New(ByVal Fecha As Date, ByVal TipoVenc As Integer)
        Me.Fecha = Fecha
        Me.TipoVencimiento = TipoVenc
    End Sub

End Class