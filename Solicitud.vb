Imports System.Data.OracleClient

Public Class Solicitud
    Implements IDisposable

    Private cn As OracleConnection
    Private WithEvents da As OracleDataAdapter  'SERREQUEST
    Private WithEvents da1 As OracleDataAdapter  'hd1clob
    Private da2 As OracleDataAdapter

    Private dt As New DataTable 'SERREQUEST
    Private dt1 As New DataTable 'hd1clob
    Private dt2 As New DataTable 'hdktask

    Private bpc As Cliente
    Private disposedValue As Boolean = False        ' Para detectar llamadas redundantes

    'NEW
    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
        Adaptadores()

        da.FillSchema(dt, SchemaType.Mapped)
        da1.FillSchema(dt1, SchemaType.Mapped)
        'da2.FillSchema(dt2, SchemaType.Mapped)

    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal Numero As String)
        Me.New(cn)
        Abrir(Numero)
    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal bpc As Cliente, ByVal cpy As Sociedad)
        Me.New(cn)
        Nueva(bpc, cpy)
    End Sub

    'SUB
    Private Sub Adaptadores()
        Dim Sql As String
        Dim sql2 As String

       
        Sql = "SELECT * FROM serrequest WHERE srenum_0 = :srenum_0"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("srenum_0", OracleType.VarChar)

        da.InsertCommand = New OracleCommandBuilder(da).GetInsertCommand
        da.UpdateCommand = New OracleCommandBuilder(da).GetUpdateCommand
        da.DeleteCommand = New OracleCommandBuilder(da).GetDeleteCommand

        'Sql = "INSERT INTO serrequest VALUES (:srenum_0, :srenumbpc_0, :sreinumbpc_0, :salfcy_0, :sredoo_0, :srebpc_0, :srebpaadd_0, :sreccn_0, :srebpcinv_0,"
        'Sql &= ":srebpcpyr_0, :srebpcgru_0, :srepjt_0, :srerep_0, :srerep_1, :srerep_2, :sremac_0, :sremacpbl_0, :stofcy_0, :srevacbpr_0, :srechgtyp_0, :sreprityp_0,"
        'Sql &= ":srecur_0, :srepte_0, :sredep_0, :cce_0, :cce_1, :cce_2, :cce_3, :cce_4, :cce_5, :cce_6, :cce_7, :cce_8, :invdtaamt_0, :invdtaamt_1, :invdtaamt_2,"
        'Sql &= ":invdtaamt_3, :invdtaamt_4, :invdtaamt_5, :invdtaamt_6, :invdtaamt_7, :invdtaamt_8, :invdtaamt_9, :invdta_0, :invdta_1, :invdta_2, :invdta_3, :invdta_4,"
        'Sql &= ":invdta_5, :invdta_6, :invdta_7, :invdta_8, :invdta_9, :typmac_0, :macflt_0, :macfltint_0, :srettr_0, :sredes_0, :typfuldes_0, :numfuldes_0, :sredesflg_0,"
        'Sql &= ":sreass_0, :sredet_0, :sreinvflg_0, :sredatass_0, :srehouass_0, :sresat_0, :tskrepcre_0, :sregralev_0, :srepiolev_0, :conspt_0, :conspttyp_0, :consptflg_0,"
        'Sql &= ":srepblgrp_0, :sretimspg_0, :timspgday_0, :timspghou_0, :timspgmnt_0, :ctimspgday_0, :ctimspghou_0, :ctimspgmnt_0, :sreesc_0, :sreesc2_0, :sreesc3_0,"
        'Sql &= ":sreresdat_0, :srereshou_0, :manupd_0, :srelok_0, :sreresren_0, :ovrcov_0, :ovrcovren_0, :ovrcovtyp_0, :srecov_0, :sretypcov_0, :srecovnum_0, :srecovaut_0,"
        'Sql &= ":srecovaus_0, :covauscla_0, :srecovctl_0, :solnum_0, :rep_0, :sreori_0, :sreoritxt_0, :sreorivcr_0, :sreorivcrl_0, :sretpl_0, :tsdcod_0, :tsdcod_1,"
        'Sql &= ":tsdcod_2, :tsdcod_3, :tsdcod_4, :trscod_0, :entcod_0, :trsfam_0, :nbrmanupd_0, :creusr_0, :credat_0, :crehou_0, :updusr_0, :upddat_0, :ysdhdeb_0,"
        'Sql &= ":ysdhfin_0, :yrptdat_0, :ypctpres_0, :ycllnum_0, :ytypsre_0, :yref_0)"
        'da.InsertCommand = New OracleCommand(Sql, cn)

        'Sql = "UPDATE serrequest SET srenumbpc_0 = :srenumbpc_0, sreinumbpc_0 = :sreinumbpc_0, salfcy_0 = :salfcy_0, sredoo_0 = :sredoo_0, srebpc_0 = :srebpc_0, "
        'Sql &= "srebpaadd_0 = :srebpaadd_0, sreccn_0 = :sreccn_0, srebpcinv_0 = :srebpcinv_0, srebpcpyr_0 = :srebpcpyr_0, srebpcgru_0 = :srebpcgru_0, srepjt_0 = :srepjt_0, "
        'Sql &= "srerep_0 = :srerep_0, srerep_1 = :srerep_1, srerep_2 = :srerep_2, sremac_0 = :sremac_0, sremacpbl_0 = :sremacpbl_0, stofcy_0 = :stofcy_0, srevacbpr_0 = :srevacbpr_0, "
        'Sql &= "srechgtyp_0 = :srechgtyp_0, sreprityp_0 = :sreprityp_0, srecur_0 = :srecur_0, srepte_0 = :srepte_0, sredep_0 = :sredep_0, cce_0 = :cce_0, cce_1 = :cce_1, cce_2 = :cce_2, "
        'Sql &= "cce_3 = :cce_3, cce_4 = :cce_4, cce_5 = :cce_5, cce_6 = :cce_6, cce_7 = :cce_7, cce_8 = :cce_8, invdtaamt_0 = :invdtaamt_0, invdtaamt_1 = :invdtaamt_1, "
        'Sql &= "invdtaamt_2 = :invdtaamt_2, invdtaamt_3 = :invdtaamt_3, invdtaamt_4 = :invdtaamt_4, invdtaamt_5 = :invdtaamt_5, invdtaamt_6 = :invdtaamt_6, invdtaamt_7 = :invdtaamt_7, "
        'Sql &= "invdtaamt_8 = :invdtaamt_8, invdtaamt_9 = :invdtaamt_9, invdta_0 = :invdta_0, invdta_1 = :invdta_1, invdta_2 = :invdta_2, invdta_3 = :invdta_3, invdta_4 = :invdta_4, "
        'Sql &= "invdta_5 = :invdta_5, invdta_6 = :invdta_6, invdta_7 = :invdta_7, invdta_8 = :invdta_8, invdta_9 = :invdta_9, typmac_0 = :typmac_0, macflt_0 = :macflt_0, macfltint_0 = :macfltint_0, "
        'Sql &= "srettr_0 = :srettr_0, sredes_0 = :sredes_0, typfuldes_0 = :typfuldes_0, numfuldes_0 = :numfuldes_0, sredesflg_0 = :sredesflg_0, sreass_0 = :sreass_0, sredet_0 = :sredet_0, "
        'Sql &= "sreinvflg_0 = :sreinvflg_0, sredatass_0 = :sredatass_0, srehouass_0 = :srehouass_0, sresat_0 = :sresat_0, tskrepcre_0 = :tskrepcre_0, sregralev_0 = :sregralev_0, "
        'Sql &= "srepiolev_0 = :srepiolev_0, conspt_0 = :conspt_0, conspttyp_0 = :conspttyp_0, consptflg_0 = :consptflg_0, srepblgrp_0 = :srepblgrp_0, sretimspg_0 = :sretimspg_0, "
        'Sql &= "timspgday_0 = :timspgday_0, timspghou_0 = :timspghou_0, timspgmnt_0 = :timspgmnt_0, ctimspgday_0 = :ctimspgday_0, ctimspghou_0 = :ctimspghou_0, ctimspgmnt_0 = :ctimspgmnt_0, "
        'Sql &= "sreesc_0 = :sreesc_0, sreesc2_0 = :sreesc2_0, sreesc3_0 = :sreesc3_0, sreresdat_0 = :sreresdat_0, srereshou_0 = :srereshou_0, manupd_0 = :manupd_0, srelok_0 = :srelok_0, "
        'Sql &= "sreresren_0 = :sreresren_0, ovrcov_0 = :ovrcov_0, ovrcovren_0 = :ovrcovren_0, ovrcovtyp_0 = :ovrcovtyp_0, srecov_0 = :srecov_0, sretypcov_0 = :sretypcov_0, "
        'Sql &= "srecovnum_0 = :srecovnum_0, srecovaut_0 = :srecovaut_0, srecovaus_0 = :srecovaus_0, covauscla_0 = :covauscla_0, srecovctl_0 = :srecovctl_0, solnum_0 = :solnum_0, "
        'Sql &= "rep_0 = :rep_0, sreori_0 = :sreori_0, sreoritxt_0 = :sreoritxt_0, sreorivcr_0 = :sreorivcr_0, sreorivcrl_0 = :sreorivcrl_0, sretpl_0 = :sretpl_0, tsdcod_0 = :tsdcod_0, "
        'Sql &= "tsdcod_1 = :tsdcod_1, tsdcod_2 = :tsdcod_2, tsdcod_3 = :tsdcod_3, tsdcod_4 = :tsdcod_4, trscod_0 = :trscod_0, entcod_0 = :entcod_0, trsfam_0 = :trsfam_0, "
        'Sql &= "nbrmanupd_0 = :nbrmanupd_0, creusr_0 = :creusr_0, credat_0 = :credat_0, crehou_0 = :crehou_0, updusr_0 = :updusr_0, upddat_0 = :upddat_0, ysdhdeb_0 = :ysdhdeb_0, "
        'Sql &= "ysdhfin_0 = :ysdhfin_0, yrptdat_0 = :yrptdat_0, ypctpres_0 = :ypctpres_0, ycllnum_0 = :ycllnum_0, ytypsre_0 = :ytypsre_0, yref_0 = :yref_0 "
        'Sql &= "WHERE srenum_0 = :srenum_0"
        'da.UpdateCommand = New OracleCommand(Sql, cn)

        'With da
        '    .SelectCommand.Parameters.Add("srenum_0", OracleType.VarChar)

        '    Parametro(.InsertCommand, "srenum_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srenumbpc_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sreinumbpc_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "salfcy_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sredoo_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srebpc_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srebpaadd_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sreccn_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srebpcinv_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srebpcpyr_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srebpcgru_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srepjt_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srerep_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srerep_1", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srerep_2", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sremac_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sremacpbl_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "stofcy_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srevacbpr_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srechgtyp_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sreprityp_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srecur_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srepte_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sredep_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "cce_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "cce_1", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "cce_2", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "cce_3", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "cce_4", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "cce_5", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "cce_6", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "cce_7", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "cce_8", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "invdtaamt_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "invdtaamt_1", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "invdtaamt_2", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "invdtaamt_3", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "invdtaamt_4", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "invdtaamt_5", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "invdtaamt_6", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "invdtaamt_7", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "invdtaamt_8", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "invdtaamt_9", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "invdta_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "invdta_1", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "invdta_2", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "invdta_3", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "invdta_4", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "invdta_5", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "invdta_6", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "invdta_7", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "invdta_8", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "invdta_9", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "typmac_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "macflt_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "macfltint_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srettr_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sredes_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "typfuldes_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "numfuldes_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sredesflg_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sreass_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sredet_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sreinvflg_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sredatass_0", OracleType.DateTime, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srehouass_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sresat_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "tskrepcre_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sregralev_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srepiolev_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "conspt_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "conspttyp_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "consptflg_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srepblgrp_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sretimspg_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "timspgday_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "timspghou_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "timspgmnt_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "ctimspgday_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "ctimspghou_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "ctimspgmnt_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sreesc_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sreesc2_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sreesc3_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sreresdat_0", OracleType.DateTime, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srereshou_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "manupd_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srelok_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sreresren_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "ovrcov_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "ovrcovren_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "ovrcovtyp_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srecov_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sretypcov_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srecovnum_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srecovaut_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srecovaus_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "covauscla_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "srecovctl_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "solnum_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "rep_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sreori_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sreoritxt_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sreorivcr_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sreorivcrl_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "sretpl_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "tsdcod_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "tsdcod_1", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "tsdcod_2", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "tsdcod_3", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "tsdcod_4", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "trscod_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "entcod_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "trsfam_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "nbrmanupd_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "creusr_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "credat_0", OracleType.DateTime, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "crehou_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "updusr_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "upddat_0", OracleType.DateTime, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "ysdhdeb_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "ysdhfin_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "yrptdat_0", OracleType.DateTime, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "ypctpres_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "ycllnum_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "ytypsre_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.InsertCommand, "yref_0", OracleType.VarChar, DataRowVersion.Current)

        '    Parametro(.UpdateCommand, "srenumbpc_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sreinumbpc_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "salfcy_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sredoo_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srebpc_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srebpaadd_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sreccn_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srebpcinv_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srebpcpyr_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srebpcgru_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srepjt_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srerep_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srerep_1", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srerep_2", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sremac_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sremacpbl_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "stofcy_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srevacbpr_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srechgtyp_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sreprityp_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srecur_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srepte_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sredep_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "cce_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "cce_1", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "cce_2", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "cce_3", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "cce_4", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "cce_5", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "cce_6", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "cce_7", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "cce_8", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "invdtaamt_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "invdtaamt_1", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "invdtaamt_2", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "invdtaamt_3", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "invdtaamt_4", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "invdtaamt_5", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "invdtaamt_6", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "invdtaamt_7", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "invdtaamt_8", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "invdtaamt_9", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "invdta_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "invdta_1", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "invdta_2", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "invdta_3", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "invdta_4", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "invdta_5", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "invdta_6", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "invdta_7", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "invdta_8", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "invdta_9", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "typmac_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "macflt_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "macfltint_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srettr_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sredes_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "typfuldes_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "numfuldes_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sredesflg_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sreass_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sredet_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sreinvflg_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sredatass_0", OracleType.DateTime, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srehouass_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sresat_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "tskrepcre_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sregralev_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srepiolev_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "conspt_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "conspttyp_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "consptflg_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srepblgrp_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sretimspg_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "timspgday_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "timspghou_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "timspgmnt_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "ctimspgday_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "ctimspghou_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "ctimspgmnt_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sreesc_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sreesc2_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sreesc3_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sreresdat_0", OracleType.DateTime, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srereshou_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "manupd_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srelok_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sreresren_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "ovrcov_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "ovrcovren_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "ovrcovtyp_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srecov_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sretypcov_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srecovnum_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srecovaut_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srecovaus_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "covauscla_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srecovctl_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "solnum_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "rep_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sreori_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sreoritxt_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sreorivcr_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sreorivcrl_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "sretpl_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "tsdcod_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "tsdcod_1", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "tsdcod_2", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "tsdcod_3", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "tsdcod_4", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "trscod_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "entcod_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "trsfam_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "nbrmanupd_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "creusr_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "credat_0", OracleType.DateTime, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "crehou_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "updusr_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "upddat_0", OracleType.DateTime, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "ysdhdeb_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "ysdhfin_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "yrptdat_0", OracleType.DateTime, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "ypctpres_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "ycllnum_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "ytypsre_0", OracleType.Number, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "yref_0", OracleType.VarChar, DataRowVersion.Current)
        '    Parametro(.UpdateCommand, "srenum_0", OracleType.VarChar, DataRowVersion.Original)
        'End With

        sql2 = "select * from hd1clob where num_0 = :num_0"
        da1 = New OracleDataAdapter(sql2, cn)

        sql2 = "INSERT INTO hd1clob VALUES(:num_0, :typ_0, :clob_0)"
        da1.InsertCommand = New OracleCommand(sql2, cn)

        sql2 = "UPDATE hd1clob SET clob_0 = :clob_0 WHERE num_0 = :num_0 AND typ_0 = :typ_0"
        da1.UpdateCommand = New OracleCommand(sql2, cn)
        With da1
            .SelectCommand.Parameters.Add("num_0", OracleType.VarChar)

            Parametro(.UpdateCommand, "clob_0", OracleType.Clob, DataRowVersion.Current)
            Parametro(.UpdateCommand, "num_0", OracleType.VarChar, DataRowVersion.Original)
            Parametro(.UpdateCommand, "typ_0", OracleType.VarChar, DataRowVersion.Original)

            Parametro(.InsertCommand, "num_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "typ_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "clob_0", OracleType.Clob, DataRowVersion.Current)
        End With

       

    End Sub
    Private Sub Modificacion()
        Dim dr As DataRow

        dr = dt.Rows(0)

        dr.BeginEdit()
        dr("upddat_0") = Date.Today
        dr("updusr_0") = USER
        dr.EndEdit()
    End Sub
    Public Sub MarcarEquiposSinTarjetas()
        Dim Sql As String

        Sql = "UPDATE sremac "
        Sql &= "SET yflgtrj_0 = 2 "
        Sql &= "WHERE srenum_0 = :srenum_0 AND yflgtrj_0 = 1 AND macnum_0 IN ( "
        Sql &= "	SELECT mac.macnum_0 "
        Sql &= "	FROM itmsales its INNER JOIN machines mac ON (its.itmref_0 = mac.macpdtcod_0) "
        Sql &= "	WHERE yflgiram_0 = 1 AND yflgsat_0 = 1 "
        Sql &= ")"

        Dim cm As New OracleCommand(Sql, cn)
        cm.Parameters.Add("srenum_0", OracleType.VarChar).Value = Numero

        Try
            cm.ExecuteNonQuery()

        Catch ex As Exception
            'MessageBox.Show(ex.Message, "MarcarEquiposSinTarjetas()", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
    Public Sub Nueva(ByVal bpc As Cliente, ByVal cpy As Sociedad)
        Dim dr As DataRow

        Me.bpc = bpc

        dt.Clear()
        dr = dt.NewRow

        dr("srenum_0") = NuevoNumero(cpy.PlantaVenta, False)
        dr("srenumbpc_0") = OrdenStr(bpc.Codigo)
        dr("sreinumbpc_0") = OrdenInt(bpc.Codigo)
        dr("salfcy_0") = cpy.PlantaVenta
        dr("sredoo_0") = " "
        dr("srebpc_0") = bpc.Codigo
        dr("srebpaadd_0") = " "
        dr("sreccn_0") = bpc.Contacto
        dr("srebpcinv_0") = bpc.Codigo
        dr("srebpcpyr_0") = bpc.TerceroPagador.Codigo
        dr("srebpcgru_0") = bpc.TerceroGrupoCodigo
        dr("srepjt_0") = " "
        dr("srerep_0") = bpc.Representante(0)
        dr("srerep_1") = bpc.Representante(1)
        dr("srerep_2") = bpc.Representante(2)
        dr("sremac_0") = " "
        dr("sremacpbl_0") = " "
        dr("stofcy_0") = cpy.PlantaStock
        dr("srevacbpr_0") = bpc.RegimenImpuesto
        dr("srechgtyp_0") = bpc.TipoCambio
        dr("sreprityp_0") = bpc.TipoPrecio
        dr("srecur_0") = bpc.Divisa
        dr("srepte_0") = bpc.CondicionPago.Codigo
        dr("sredep_0") = " "
        dr("cce_0") = " "
        dr("cce_1") = " "
        dr("cce_2") = " "
        dr("cce_3") = " "
        dr("cce_4") = " "
        dr("cce_5") = " "
        dr("cce_6") = " "
        dr("cce_7") = " "
        dr("cce_8") = " "
        dr("invdtaamt_0") = 0
        dr("invdtaamt_1") = 0
        dr("invdtaamt_2") = 0
        dr("invdtaamt_3") = 0
        dr("invdtaamt_4") = 0
        dr("invdtaamt_5") = 0
        dr("invdtaamt_6") = 0
        dr("invdtaamt_7") = 0
        dr("invdtaamt_8") = 0
        dr("invdtaamt_9") = 0
        dr("invdta_0") = 10
        dr("invdta_1") = 0
        dr("invdta_2") = 0
        dr("invdta_3") = 0
        dr("invdta_4") = 0
        dr("invdta_5") = 0
        dr("invdta_6") = 0
        dr("invdta_7") = 0
        dr("invdta_8") = 0
        dr("invdta_9") = 0
        dr("typmac_0") = 1
        dr("macflt_0") = " "
        dr("macfltint_0") = 2
        dr("srettr_0") = " "
        dr("sredes_0") = " "
        dr("typfuldes_0") = "SRE"
        dr("numfuldes_0") = dr("srenum_0", DataRowVersion.Proposed)
        dr("sredesflg_0") = 0
        dr("sreass_0") = 2
        dr("sredet_0") = USER
        dr("sreinvflg_0") = 1
        dr("sredatass_0") = #12/31/1599#
        dr("srehouass_0") = " "
        dr("sresat_0") = " "
        dr("tskrepcre_0") = 1
        dr("sregralev_0") = "D1"
        dr("srepiolev_0") = "B1"
        dr("conspt_0") = " "
        dr("conspttyp_0") = " "
        dr("consptflg_0") = 0
        dr("srepblgrp_0") = " "
        dr("sretimspg_0") = 0
        dr("timspgday_0") = 0
        dr("timspghou_0") = 0
        dr("timspgmnt_0") = 0
        dr("ctimspgday_0") = 0
        dr("ctimspghou_0") = 0
        dr("ctimspgmnt_0") = 0
        dr("sreesc_0") = 0
        dr("sreesc2_0") = 0
        dr("sreesc3_0") = 0
        dr("sreresdat_0") = Date.Today.AddDays(7)
        dr("srereshou_0") = Date.Now.ToString("ddmm")
        dr("manupd_0") = 1
        dr("srelok_0") = 1
        dr("sreresren_0") = " "
        dr("ovrcov_0") = 3
        dr("ovrcovren_0") = 0
        dr("ovrcovtyp_0") = " "
        dr("srecov_0") = 0
        dr("sretypcov_0") = 0
        dr("srecovnum_0") = " "
        dr("srecovaut_0") = 1
        dr("srecovaus_0") = " "
        dr("covauscla_0") = " "
        dr("srecovctl_0") = " "
        dr("solnum_0") = " "
        dr("rep_0") = " "
        dr("sreori_0") = 1
        dr("sreoritxt_0") = "Creacion manual"
        dr("sreorivcr_0") = " "
        dr("sreorivcrl_0") = 0
        dr("sretpl_0") = " "
        dr("tsdcod_0") = " "
        dr("tsdcod_1") = " "
        dr("tsdcod_2") = " "
        dr("tsdcod_3") = " "
        dr("tsdcod_4") = " "
        dr("trscod_0") = " "
        dr("entcod_0") = " "
        dr("trsfam_0") = " "
        dr("nbrmanupd_0") = 0
        dr("creusr_0") = USER
        dr("credat_0") = Date.Today
        dr("crehou_0") = Date.Now.ToString("ddmm")
        dr("updusr_0") = " "
        dr("upddat_0") = #12/31/1599#
        dr("ysdhdeb_0") = " "
        dr("ysdhfin_0") = " "
        dr("yrptdat_0") = #12/31/1599#
        dr("ypctpres_0") = 0
        dr("ycllnum_0") = " "
        dr("ytypsre_0") = 2
        dr("yref_0") = " "
        dt.Rows.Add(dr)

        Grabar()

    End Sub
    Public Sub Grabar()
        da.Update(dt)
        da1.Update(dt1)

    End Sub
    Public Sub CerrarSolicitud()
        Dim dr As DataRow = dt.Rows(0)

        dr.BeginEdit()
        dr("sreass_0") = EstadoSolicitud.Cerrada
        dr("sredet_0") = " "
        dr("sredatass_0") = Date.Today
        dr("srehouass_0") = Date.Now.ToString("ddMM")
        dr("rep_0") = USER
        dr("nbrmanupd_0") = 1
        dr("updusr_0") = USER
        dr("upddat_0") = Date.Today
        dr.EndEdit()

    End Sub
    Public Sub LimpiarParque()
        Dim Sql As String = "DELETE FROM sremac WHERE srenum_0 = :srenum_0 AND macnum_0 NOT IN (SELECT macnum_0 FROM machines)"
        Dim cm As New OracleCommand(Sql, cn)
        cm.Parameters.Add("srenum_0", OracleType.VarChar).Value = Numero

        Try
            cm.ExecuteNonQuery()

        Catch ex As Exception
            'MessageBox.Show(ex.Message, "Solicitud.LimpiarParque()", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

        cm.Dispose()

    End Sub
   
    'FUNCTION
    Private Function NuevoNumero(ByVal Planta As String, Optional ByVal ModoTest As Boolean = False) As String
        Dim dr As OracleDataReader
        Dim Serie As Integer
        Dim cm1 As New OracleCommand("SELECT valeur_0 FROM avalnum WHERE codnum_0 = 'SRE'", cn)
        Dim cm2 As New OracleCommand("UPDATE avalnum SET valeur_0 = :valeur_0 WHERE codnum_0 = 'SRE'", cn)
        cm2.Parameters.Add("valeur_0", OracleType.Number)

        dr = cm1.ExecuteReader(CommandBehavior.SingleResult)

        If dr.Read() Then
            Serie = CType(dr("valeur_0"), Integer)

            'Aumento el numerador
            cm2.Parameters("valeur_0").Value = Serie + 1

            'Si es test no adelanta el numerador en la tabla de contadores
            If Not ModoTest Then cm2.ExecuteNonQuery()
        End If
        dr.Close()

        cm1.Dispose()
        cm2.Dispose()
        dr.Dispose()

        Dim s As String = String.Format("000000000000{0}", Serie)

        Return Planta & s.Substring(s.Length - 12)

    End Function
    Private Function OrdenInt(ByVal Cliente As String) As Integer
        Dim dr As OracleDataReader
        Dim cm As New OracleCommand("SELECT MAX(sreinumbpc_0) FROM serrequest WHERE srebpc_0 = :srebpc_0", cn)
        cm.Parameters.Add("srebpc_0", OracleType.VarChar).Value = Cliente

        dr = cm.ExecuteReader

        If dr.Read Then
            If IsDBNull(dr(0)) Then
                OrdenInt = 1

            Else
                OrdenInt = CInt(dr(0)) + 1

            End If

        Else
            OrdenInt = 1

        End If
        dr.Close()

        cm.Dispose()

    End Function
    Private Function OrdenStr(ByVal Cliente As String) As String
        Dim dr As OracleDataReader
        Dim cm As New OracleCommand("SELECT MAX(to_number(srenumbpc_0)) FROM serrequest WHERE srebpc_0 = :srebpc_0", cn)
        cm.Parameters.Add("srebpc_0", OracleType.VarChar).Value = Cliente

        dr = cm.ExecuteReader

        If dr.Read Then
            If IsDBNull(dr(0)) Then
                OrdenStr = "1"

            Else
                OrdenStr = CStr(CInt(dr(0)) + 1)

            End If

        Else
            OrdenStr = "1"

        End If
        dr.Close()

        cm.Dispose()

    End Function
    Public Function Abrir(ByVal Numero As String) As Boolean
        Dim dr As DataRow

        dt.Clear()
        dt1.Clear()
        da.SelectCommand.Parameters("srenum_0").Value = Numero
        da1.SelectCommand.Parameters("num_0").Value = Numero
        da.Fill(dt) 'serrequest
        da1.Fill(dt1) 'hd1clob

        If dt.Rows.Count = 1 Then
            dr = dt.Rows(0)
            Return True
        Else
            Return False
        End If

    End Function
    Public Function CrearIntervencion(ByVal Sucursal As String, ByVal TipoIntervencion As String, Optional ByVal diadesde As String = "", Optional ByVal diahasta As String = "", Optional ByVal tardedesde As String = "", Optional ByVal tardehasta As String = "") As Intervencion
        Dim itn As New Intervencion(cn)

        itn.Nueva(Me, TipoIntervencion, Sucursal, diadesde, diahasta, tardedesde, tardehasta)
        itn.Referencia = Me.Referencia

        Return itn

    End Function

    'EVENTS
    'Private Sub da_RowUpdated(ByVal sender As Object, ByVal e As System.Data.OracleClient.OracleRowUpdatedEventArgs) Handles da.RowUpdated

    '    If e.Status = UpdateStatus.ErrorsOccurred Then Exit Sub

    '    Select Case e.StatementType
    '        Case StatementType.Insert
    '            Auditoria(cn, "SERREQUEST", 1, e.Row("srenum_0").ToString)

    '        Case StatementType.Update
    '            Auditoria(cn, "SERREQUEST", 2, e.Row("srenum_0").ToString)

    '        Case StatementType.Delete
    '            Auditoria(cn, "SERREQUEST", 3, e.Row("srenum_0").ToString)

    '    End Select

    'End Sub
    'PROPERTY
    Public ReadOnly Property Numero() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("srenum_0").ToString
        End Get
    End Property
    Public ReadOnly Property PlantaVenta() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("salfcy_0").ToString
        End Get
    End Property
    Public ReadOnly Property PlantaAlmacenamiento() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("stofcy_0").ToString
        End Get
    End Property
    Public ReadOnly Property Tercero() As Cliente
        Get
            Dim dr As DataRow = dt.Rows(0)
            Dim bpc As New Cliente(cn)
            bpc.Abrir(dr("srebpc_0").ToString)
            Return bpc
        End Get
    End Property
    Public ReadOnly Property Estado() As EstadoSolicitud
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CType(dr("sreass_0"), EstadoSolicitud)
        End Get
    End Property
    Public Property Contacto() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("sreccn_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("sreccn_0") = IIf(value.Trim = "", " ", value.Trim)
            dr.EndEdit()
        End Set
    End Property
    Public Property Referencia() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("yref_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("yref_0") = IIf(value.Trim = "", " ", value.Trim)
            dr.EndEdit()
        End Set
    End Property
    Public Property updusr() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("updusr_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("updusr_0") = IIf(value.Trim = "", " ", value.Trim)
            dr.EndEdit()
        End Set
    End Property
    Public Property descripcion() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("sredes_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("sredes_0") = IIf(value.Trim = "", " ", value.Trim)
            dr.EndEdit()
        End Set
    End Property
    Public Property estado_facturacion() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CInt(dr("sreinvflg_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("sreinvflg_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property descripcion_full() As String
        Get
            Dim txt As String = ""
            If dt1.Rows.Count = 1 Then txt = dt1.Rows(0).Item("clob_0").ToString
            Return txt
        End Get
        Set(ByVal value As String)
            Dim dr1 As DataRow '= dt1.Rows(0)
            If dt1.Rows.Count = 0 Then
                dr1 = dt1.NewRow
                dr1("num_0") = Me.Numero
                dr1("typ_0") = "SRE"
                dr1("clob_0") = value
                dt1.Rows.Add(dr1)
            Else
                dr1 = dt1.Rows(0)
                If dr1.RowState = DataRowState.Deleted Then dr1.RejectChanges()
                dr1.BeginEdit()
                dr1("clob_0") = value
                dr1.EndEdit()
            End If
        End Set
    End Property


    Public ReadOnly Property Facturacion() As FlagFacturacion
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CType(dr("sreinvflg_0"), FlagFacturacion)
        End Get
    End Property
    Public ReadOnly Property Vendedor() As Vendedor
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return New Vendedor(cn, dr("srerep_0").ToString)
        End Get
    End Property
    Public ReadOnly Property CantidadIntervenciones() As Integer
        Get
            Dim Sql As String = "SELECT num_0 FROM interven WHERE srvdemnum_0 = :srvdemnum_0"
            Dim da As OracleDataAdapter
            Dim dt As New DataTable

            da = New OracleDataAdapter(Sql, cn)
            da.SelectCommand.Parameters.Add("srvdemnum_0", OracleType.VarChar).Value = Numero
            da.Fill(dt)
            da.Dispose()

            Return dt.Rows.Count
        End Get
    End Property
    Public ReadOnly Property CondicionPago() As CondicionPago
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return New CondicionPago(cn, dr("srepte_0").ToString)
        End Get
    End Property
    Public Property Contrato() As ContratoServicio
        Get
            Dim dr As DataRow
            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return New ContratoServicio(cn, dr("conspt_0").ToString)

            Else
                Return Nothing

            End If
        End Get
        Set(ByVal value As ContratoServicio)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("conspt_0") = value.Numero
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property TieneContrato() As Boolean
        Get
            Dim dr As DataRow
            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return (dr("conspt_0").ToString <> " ")
            Else
                Return False
            End If
        End Get
    End Property

    Public WriteOnly Property ContratoTipo() As String
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("conspttyp_0") = IIf(value.Trim = "", " ", value.Trim)
            dr.EndEdit()
        End Set
    End Property
    Public WriteOnly Property CoberturaGlobal() As Integer
        Set(ByVal value As Integer)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("ovrcov_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property CantidadConsumos() As Integer
        Get
            If dt2 Is Nothing Then
                da2.SelectCommand.Parameters("srenum").Value = Me.Numero
                Try
                    dt2 = New DataTable
                    da2.Fill(dt2)
                Catch ex As Exception
                End Try
            End If

            Return dt2.Rows.Count
        End Get
    End Property

    Public Property TipoCovertura() As Integer
        Get
            Dim dr As DataRow
            Dim i As Integer = 0

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                i = CInt(dr("ovrcov_0"))
            End If

            Return i
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("ovrcov_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: Liberar otro estado (objetos administrados).
            End If

            ' TODO: Liberar su propio estado (objetos no administrados).
            ' TODO: Establecer campos grandes como Null.
            da.Dispose()
            dt.Dispose()

        End If
        Me.disposedValue = True
    End Sub

    ' Visual Basic agregó este código para implementar correctamente el modelo descartable.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' No cambie este código. Coloque el código de limpieza en Dispose (ByVal que se dispone como Boolean).
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
        Dispose(False)
    End Sub
    '======================================================================================
    'ENUMERACIONES
    '======================================================================================
    Public Enum EstadoSolicitud
        SinEstado
        Dispatching
        Colaborador
        Cola
        ServicioComer
        Cerrada
    End Enum
    Public Enum FlagFacturacion
        SinEstado
        NoFacturable
        Facturable
        Facturada
    End Enum

End Class