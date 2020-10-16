Imports System.Data.OracleClient
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine

Public Class Intervencion
    Implements IDisposable, IRuteable

    Private cn As OracleConnection

    Private l_Equipos As Integer = 0
    Private l_Mangas As Integer = 0
    Private l_Peso As Double = 0
    Private l_Peso2 As Double = 0 'Peso para unigis sin prestamos
    Private l_EsTarea As Boolean = False

    Private l_PrestamosExt As Integer = 0
    Private l_PrestamosMan As Integer = 0
    Private l_RechazosExt As Integer = 0
    Private l_RechazosMan As Integer = 0
    Private l_TieneCarro As Boolean = False
    Private l_Varios As Boolean = False

    Private l_Cliente As Cliente = Nothing
    Private l_Sucursal As Sucursal = Nothing

    Private WithEvents da1 As OracleDataAdapter 'Intervencion (INTERVEN)
    Private WithEvents da2 As OracleDataAdapter 'Detalle (YITNDET)
    Private WithEvents da3 As OracleDataAdapter 'HD6Clob
    Private da4 As OracleDataAdapter 'aclob
    Private da5 As OracleDataAdapter 'hdktask
    Private da6 As OracleDataAdapter 'hdktaskinv
    Private da7 As OracleDataAdapter

    Private dt1 As New DataTable    'Intervencion (INTERVEN)
    Private dt2 As New DataTable    'Detalle      (YITNDET)
    Private dt3 As New DataTable    'HD6Clob
    Private dt4 As New DataTable    'aclob
    Private dt5 As New DataTable    'hdktask
    Private dt6 As New DataTable    'hdktaskinv
    Private dt7 As New DataTable    'gaccdudate (Parte de Cobranzas)

    Private disposedValue As Boolean = False        ' Para detectar llamadas redundantes
    Public Const DB_USR As String = "GEOPROD"
    Public Const DB_PWD As String = "tiger"

    'EVENTS
    Public Event ParquesMarcados(ByVal sender As Object, ByVal e As ParquesEvenArgs)
    Public Event IntervencionAbierta(ByVal sender As Object)

    'NEW
    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
        Adaptadores()

        da1.FillSchema(dt1, SchemaType.Mapped)
        da2.FillSchema(dt2, SchemaType.Mapped)
        da3.FillSchema(dt3, SchemaType.Mapped)
        da4.FillSchema(dt4, SchemaType.Mapped)
        da5.FillSchema(dt5, SchemaType.Mapped)
        da6.FillSchema(dt6, SchemaType.Mapped)

        CalcularIdentidad()

    End Sub

    'SUB
    Public Sub Grabar()
        'Si la intervención es nueva. Pido numero de intervención
        da2.Update(dt2) 'yitndet
        da3.Update(dt3) 'hd6clob
        da1.Update(dt1) 'interven
        da4.Update(dt4) 'aclob
        da5.Update(dt5) 'hdktask
        da6.Update(dt6) 'hdktaskinv
    End Sub
    Private Sub Adaptadores()
        Dim Sql As String

        Sql = "SELECT * FROM interven WHERE num_0 = :num_0"
        da1 = New OracleDataAdapter(Sql, cn)

        Sql = "INSERT INTO interven VALUES (:num_0, :salfcy_0, :srvdemnum_0, :bpc_0, :ccn_0, :mac_0, :macgru_0, :typ_0, :dat_0, :datend_0, :hou_0, :houend_0, :fulday_0, :wee_0, :dur_0, :datx_0, :dtckil_0, "
        Sql &= ":tritim_0, :obj_0, :typfulobj_0, :numfulobj_0, :objflg_0, :rep_0, :timspg_0, :htottimspg_0, :mtottimspg_0, :mantimflg_0, :pblsol_0, :srvconcov_0, :connum_0, "
        Sql &= ":ordnum_0, :sco_0, :sconum_0, :scoamt_0, :cur_0, :bpaadd_0, :drn_0, :itntypadd_0, :itncodadd_0, :itnrecadd_0, :add_0, :add_1, :add_2, :zip_0, :cty_0, "
        Sql &= ":cry_0, :sat_0, :tel_0, :xportero_0, :xpor_tel_0, :xmailfc_0, :iffadd_0, :rer_0, :rer_1, :rer_2, :rer_3, :rer_4, :rer_5, :rer_6, :rer_7, :rer_8, :rer_9, :rer_10, :rer_11, :rer_12, :rer_13, "
        Sql &= ":rer_14, :don_0, :typfulrpo_0, :numfulrpo_0, :rpo_0, :rpoflg_0, :itnori_0, :itnoritxt_0, :itnorivcr_0, :itnorivcrl_0, :creusr_0, :credat_0, :crehou_0, "
        Sql &= ":updusr_0, :upddat_0, :ypctpres_0, :ycllnum_0, :zflgtrip_0, :yflgsdh_0, :yhdtamtinv_0, :ymrkitn_0, :tripnum_0, :ysdhdeb_0, :ysdhfin_0, :mdl_0, :yref_0, "
        Sql &= ":dlvpio_0, :yflgret_0, :yhdesde1_0, :yhhasta1_0, :yhdesde2_0, :yhhasta2_0, :yotr_0, :xauto_0, "
        Sql &= ":ymesa_0, :yobserva_0, :xnoconf_0, :xsector_0, :xweb_0, :xtrans_0, :xgeo_0, :xrpt_0, :xcer_0, :xtanda_0,:ocf_0, :yobsrec_0, :yobsitn_0, "
        Sql &= ":xcarzon_0, :xcardat_0, :xcartyp_0)"
        da1.InsertCommand = New OracleCommand(Sql, cn)

        Sql = "UPDATE interven SET salfcy_0=:salfcy_0, srvdemnum_0=:srvdemnum_0, bpc_0=:bpc_0, ccn_0=:ccn_0, mac_0=:mac_0, macgru_0=:macgru_0, typ_0=:typ_0, dat_0=:dat_0, "
        Sql &= "datend_0=:datend_0, hou_0=:hou_0, houend_0=:houend_0, fulday_0=:fulday_0, wee_0=:wee_0, dur_0=:dur_0, datx_0=:datx_0, dtckil_0=:dtckil_0, tritim_0=:tritim_0, "
        Sql &= "obj_0=:obj_0, typfulobj_0=:typfulobj_0, numfulobj_0=:numfulobj_0, objflg_0=:objflg_0, rep_0=:rep_0, timspg_0=:timspg_0, htottimspg_0=:htottimspg_0, "
        Sql &= "mtottimspg_0=:mtottimspg_0, mantimflg_0=:mantimflg_0, pblsol_0=:pblsol_0, srvconcov_0=:srvconcov_0, connum_0=:connum_0, ordnum_0=:ordnum_0, sco_0=:sco_0, "
        Sql &= "sconum_0=:sconum_0, scoamt_0=:scoamt_0, cur_0=:cur_0, bpaadd_0=:bpaadd_0, drn_0=:drn_0, itntypadd_0=:itntypadd_0, itncodadd_0=:itncodadd_0, itnrecadd_0=:itnrecadd_0, "
        Sql &= "add_0=:add_0, add_1=:add_1, add_2=:add_2, zip_0=:zip_0, cty_0=:cty_0, cry_0=:cry_0, sat_0=:sat_0, tel_0=:tel_0, xportero_0=:xportero_0, xpor_tel_0=:xpor_tel_0, xmailfc_0=:xmailfc_0, iffadd_0=:iffadd_0, rer_0=:rer_0, rer_1=:rer_1, "
        Sql &= "rer_2=:rer_2, rer_3=:rer_3, rer_4=:rer_4, rer_5=:rer_5, rer_6=:rer_6, rer_7=:rer_7, rer_8=:rer_8, rer_9=:rer_9, rer_10=:rer_10, rer_11=:rer_11, rer_12=:rer_12, "
        Sql &= "rer_13=:rer_13, rer_14=:rer_14, don_0=:don_0, typfulrpo_0=:typfulrpo_0, numfulrpo_0=:numfulrpo_0, rpo_0=:rpo_0, rpoflg_0=:rpoflg_0, itnori_0=:itnori_0, itnoritxt_0=:itnoritxt_0, "
        Sql &= "itnorivcr_0=:itnorivcr_0, itnorivcrl_0=:itnorivcrl_0, creusr_0=:creusr_0, credat_0=:credat_0, crehou_0=:crehou_0, updusr_0=:updusr_0, upddat_0=:upddat_0, "
        Sql &= "ypctpres_0=:ypctpres_0, ycllnum_0=:ycllnum_0, zflgtrip_0=:zflgtrip_0, yflgsdh_0=:yflgsdh_0, yhdtamtinv_0=:yhdtamtinv_0, ymrkitn_0=:ymrkitn_0, "
        Sql &= "tripnum_0=:tripnum_0, ysdhdeb_0=:ysdhdeb_0, ysdhfin_0=:ysdhfin_0, mdl_0=:mdl_0, yref_0=:yref_0, dlvpio_0=:dlvpio_0, yflgret_0=:yflgret_0, yhdesde1_0=:yhdesde1_0, "
        Sql &= "yhhasta1_0=:yhhasta1_0, yhdesde2_0=:yhdesde2_0, yhhasta2_0=:yhhasta2_0, yotr_0=:yotr_0, xauto_0=:xauto_0, "
        Sql &= "ymesa_0=:ymesa_0, yobserva_0=:yobserva_0, yobsrec_0=:yobsrec_0, yobsitn_0=:yobsitn_0, xnoconf_0=:xnoconf_0, xsector_0 = :xsector_0, xweb_0 = :xweb_0, xtrans_0 = :xtrans_0, xgeo_0 = :xgeo_0, xrpt_0 = :xrpt_0, xcer_0 = :xcer_0, xtanda_0=:xtanda_0, ocf_0 = :ocf_0, "
        Sql &= "xcarzon_0=:xcarzon_0, xcardat_0=:xcardat_0, xcartyp_0=:xcartyp_0 "
        Sql &= "WHERE num_0 = :num_0"
        da1.UpdateCommand = New OracleCommand(Sql, cn)
        With da1
            .SelectCommand.Parameters.Add("num_0", OracleType.VarChar)

            Parametro(.InsertCommand, "num_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "salfcy_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "srvdemnum_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "bpc_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "ccn_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "mac_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "macgru_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "typ_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "dat_0", OracleType.DateTime, DataRowVersion.Current)
            Parametro(.InsertCommand, "datend_0", OracleType.DateTime, DataRowVersion.Current)
            Parametro(.InsertCommand, "hou_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "houend_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "fulday_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "wee_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "dur_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "datx_0", OracleType.DateTime, DataRowVersion.Current)
            Parametro(.InsertCommand, "dtckil_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "tritim_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "obj_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "typfulobj_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "numfulobj_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "objflg_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "rep_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "timspg_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "htottimspg_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "mtottimspg_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "mantimflg_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "pblsol_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "srvconcov_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "connum_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "ordnum_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "sco_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "sconum_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "scoamt_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "cur_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "bpaadd_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "drn_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "itntypadd_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "itncodadd_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "itnrecadd_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "add_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "add_1", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "add_2", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "zip_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "cty_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "cry_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "sat_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "tel_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "xportero_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "xpor_tel_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "xmailfc_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "iffadd_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "rer_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "rer_1", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "rer_2", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "rer_3", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "rer_4", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "rer_5", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "rer_6", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "rer_7", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "rer_8", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "rer_9", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "rer_10", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "rer_11", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "rer_12", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "rer_13", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "rer_14", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "don_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "typfulrpo_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "numfulrpo_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "rpo_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "rpoflg_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "itnori_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "itnoritxt_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "itnorivcr_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "itnorivcrl_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "creusr_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "credat_0", OracleType.DateTime, DataRowVersion.Current)
            Parametro(.InsertCommand, "crehou_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "updusr_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "upddat_0", OracleType.DateTime, DataRowVersion.Current)
            Parametro(.InsertCommand, "ypctpres_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "ycllnum_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "zflgtrip_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "yflgsdh_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "yhdtamtinv_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "ymrkitn_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "tripnum_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "ysdhdeb_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "ysdhfin_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "mdl_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "yref_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "dlvpio_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "yflgret_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "yhdesde1_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "yhhasta1_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "yhdesde2_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "yhhasta2_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "yotr_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "xauto_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "ymesa_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "yobserva_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "xnoconf_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "xsector_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "xweb_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "xtrans_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "xgeo_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "xrpt_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "xcer_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "xtanda_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "ocf_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "yobsrec_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "yobsitn_0", OracleType.VarChar, DataRowVersion.Current)

            Parametro(.InsertCommand, "xcarzon_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "xcardat_0", OracleType.DateTime, DataRowVersion.Current)
            Parametro(.InsertCommand, "xcartyp_0", OracleType.Number, DataRowVersion.Current)

            Parametro(.UpdateCommand, "salfcy_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "srvdemnum_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "bpc_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "ccn_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "mac_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "macgru_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "typ_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "dat_0", OracleType.DateTime, DataRowVersion.Current)
            Parametro(.UpdateCommand, "datend_0", OracleType.DateTime, DataRowVersion.Current)
            Parametro(.UpdateCommand, "hou_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "houend_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "fulday_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "wee_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "dur_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "datx_0", OracleType.DateTime, DataRowVersion.Current)
            Parametro(.UpdateCommand, "dtckil_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "tritim_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "obj_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "typfulobj_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "numfulobj_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "objflg_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "rep_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "timspg_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "htottimspg_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "mtottimspg_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "mantimflg_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "pblsol_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "srvconcov_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "connum_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "ordnum_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "sco_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "sconum_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "scoamt_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "cur_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "bpaadd_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "drn_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "itntypadd_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "itncodadd_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "itnrecadd_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "add_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "add_1", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "add_2", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "zip_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "cty_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "cry_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "sat_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "tel_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "xportero_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "xpor_tel_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "xmailfc_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "iffadd_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "rer_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "rer_1", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "rer_2", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "rer_3", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "rer_4", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "rer_5", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "rer_6", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "rer_7", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "rer_8", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "rer_9", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "rer_10", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "rer_11", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "rer_12", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "rer_13", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "rer_14", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "don_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "typfulrpo_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "numfulrpo_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "rpo_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "rpoflg_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "itnori_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "itnoritxt_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "itnorivcr_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "itnorivcrl_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "creusr_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "credat_0", OracleType.DateTime, DataRowVersion.Current)
            Parametro(.UpdateCommand, "crehou_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "updusr_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "upddat_0", OracleType.DateTime, DataRowVersion.Current)
            Parametro(.UpdateCommand, "ypctpres_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "ycllnum_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "zflgtrip_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "yflgsdh_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "yhdtamtinv_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "ymrkitn_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "tripnum_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "ysdhdeb_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "ysdhfin_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "mdl_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "yref_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "dlvpio_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "yflgret_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "yhdesde1_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "yhhasta1_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "yhdesde2_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "yhhasta2_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "yotr_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "xauto_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "ymesa_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "yobserva_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "xnoconf_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "xsector_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "xweb_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "xtrans_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "xgeo_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "xrpt_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "xcer_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "xtanda_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "num_0", OracleType.VarChar, DataRowVersion.Original)
            Parametro(.UpdateCommand, "ocf_0", OracleType.VarChar, DataRowVersion.Original)
            Parametro(.UpdateCommand, "yobsrec_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "yobsitn_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "xcarzon_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "xcardat_0", OracleType.DateTime, DataRowVersion.Current)
            Parametro(.UpdateCommand, "xcartyp_0", OracleType.Number, DataRowVersion.Current)
        End With

        'TABLA ITNDET
        Sql = "SELECT * FROM yitndet WHERE num_0 = :num_0"
        da2 = New OracleDataAdapter(Sql, cn)

        Sql = "INSERT INTO yitndet VALUES(:numlig_0, :num_0, :itmref_0, :qty_0, :uom_0, :tqty_0, :tuom_0, :yqty2_0, :factura_0, :typlig_0, :amt_0, :srenum_0)"
        da2.InsertCommand = New OracleCommand(Sql, cn)

        Sql = "UPDATE yitndet SET itmref_0=:itmref_0, qty_0=:qty_0, uom_0=:uom_0, tqty_0=:tqty_0, tuom_0=:tuom_0, yqty2_0=:yqty2_0, factura_0=:factura_0, typlig_0=:typlig_0, amt_0=:amt_0, srenum_0 = :srenum_0 WHERE num_0=:num_0 AND numlig_0=:numlig_0"
        da2.UpdateCommand = New OracleCommand(Sql, cn)

        Sql = "DELETE FROM yitndet WHERE num_0 = :num_0 AND numlig_0 = :numlig_0"
        da2.DeleteCommand = New OracleCommand(Sql, cn)

        With da2
            .SelectCommand.Parameters.Add("num_0", OracleType.VarChar)

            Parametro(.InsertCommand, "numlig_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "num_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "itmref_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "qty_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "uom_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "tqty_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "tuom_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "yqty2_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "factura_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "typlig_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "amt_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "srenum_0", OracleType.VarChar, DataRowVersion.Current)

            Parametro(.UpdateCommand, "itmref_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "qty_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "uom_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "tqty_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "tuom_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "yqty2_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "factura_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "typlig_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "amt_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "srenum_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "num_0", OracleType.VarChar, DataRowVersion.Original)
            Parametro(.UpdateCommand, "numlig_0", OracleType.Number, DataRowVersion.Original)

            Parametro(.DeleteCommand, "num_0", OracleType.VarChar, DataRowVersion.Original)
            Parametro(.DeleteCommand, "numlig_0", OracleType.Number, DataRowVersion.Original)
        End With

        'TABLA HD6CLOB
        Sql = "SELECT * FROM hd6clob WHERE num_0 = :num_0 AND typ_0 = 'ITNOBJ'"
        da3 = New OracleDataAdapter(Sql, cn)

        Sql = "UPDATE hd6clob SET clob_0 = :clob_0 WHERE num_0 = :num_0 AND typ_0 = :typ_0"
        da3.UpdateCommand = New OracleCommand(Sql, cn)

        Sql = "INSERT INTO hd6clob VALUES(:num_0, :typ_0, :clob_0)"
        da3.InsertCommand = New OracleCommand(Sql, cn)

        Sql = "DELETE FROM hd6clob WHERE num_0 = :num_0 AND typ_0 = :typ_0"
        da3.DeleteCommand = New OracleCommand(Sql, cn)

        With da3
            .SelectCommand.Parameters.Add("num_0", OracleType.VarChar)

            Parametro(.UpdateCommand, "clob_0", OracleType.Clob, DataRowVersion.Current)
            Parametro(.UpdateCommand, "num_0", OracleType.VarChar, DataRowVersion.Original)
            Parametro(.UpdateCommand, "typ_0", OracleType.VarChar, DataRowVersion.Original)

            Parametro(.InsertCommand, "num_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "typ_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "clob_0", OracleType.Clob, DataRowVersion.Current)

            Parametro(.DeleteCommand, "num_0", OracleType.VarChar, DataRowVersion.Original)
            Parametro(.DeleteCommand, "typ_0", OracleType.VarChar, DataRowVersion.Original)

        End With

        'tabla aclob
        Sql = "select * from aclob where ident1_0 = :ident1_0 "
        da4 = New OracleDataAdapter(Sql, cn)
        da4.SelectCommand.Parameters.Add("ident1_0", OracleType.VarChar)
        da4.InsertCommand = New OracleCommandBuilder(da4).GetInsertCommand
        da4.UpdateCommand = New OracleCommandBuilder(da4).GetUpdateCommand
        da4.DeleteCommand = New OracleCommandBuilder(da4).GetDeleteCommand

        Sql = "select * from hdktask where itnnum_0 = :itnnum_0 "
        da5 = New OracleDataAdapter(Sql, cn)
        da5.SelectCommand.Parameters.Add("itnnum_0", OracleType.VarChar)
        da5.InsertCommand = New OracleCommandBuilder(da5).GetInsertCommand
        da5.UpdateCommand = New OracleCommandBuilder(da5).GetUpdateCommand
        da5.DeleteCommand = New OracleCommandBuilder(da5).GetDeleteCommand

        Sql = "select * from hdktaskinv where srenum_0 = :srenum_0 "
        da6 = New OracleDataAdapter(Sql, cn)
        da6.SelectCommand.Parameters.Add("srenum_0", OracleType.VarChar)
        da6.InsertCommand = New OracleCommandBuilder(da6).GetInsertCommand


        'Consulta para Parte de Cobranza
        Sql = "select gac.*"
        Sql &= "from sinvoicev sih inner join gaccdudate gac on (sih.num_0 = gac.num_0) "
        Sql &= "where sihorinum_0 = :num"
        da7 = New OracleDataAdapter(Sql, cn)
        da7.SelectCommand.Parameters.Add("num", OracleType.VarChar)

    End Sub
    Private Sub BuscarEquiposMarcados()
        Dim da As New OracleDataAdapter("SELECT macnum_0 FROM machines WHERE xitn_0 = :xitn_0", cn)
        Dim dt As New DataTable

        da.SelectCommand.Parameters.Add("xitn_0", OracleType.VarChar).Value = Numero

        da.Fill(dt)

        'Disparo evento que devuelve los parques ocultados por la intervencion
        If dt.Rows.Count > 0 Then RaiseEvent ParquesMarcados(Me, New ParquesEvenArgs(dt))

        da.Dispose()

    End Sub
    Private Sub CalcularIdentidad()
        Dim InicioIdentidad As Integer = 0
        Dim dr As DataRow

        For Each dr In dt2.Rows
            If CInt(dr("numlig_0")) > InicioIdentidad Then InicioIdentidad = CInt(dr("numlig_0"))
        Next

        InicioIdentidad += 1000

        With dt2.Columns
            .Item("numlig_0").AutoIncrement = True
            .Item("numlig_0").AutoIncrementSeed = InicioIdentidad
            .Item("numlig_0").AutoIncrementStep = 1000

            .Item("qty_0").DefaultValue = 0
            .Item("uom_0").DefaultValue = "UN"
            .Item("tqty_0").DefaultValue = 0
            .Item("tuom_0").DefaultValue = "UN"
            .Item("yqty2_0").DefaultValue = 0
            .Item("factura_0").DefaultValue = 2
            .Item("typlig_0").DefaultValue = 1
            .Item("amt_0").DefaultValue = 0
        End With

    End Sub
    Public Sub AgregarDetalle(ByVal dt As DataTable)
        Dim dr As DataRow
        Dim dr2 As DataRow
        Dim i As Integer

        'Se elimina el contenido actual de la tabla detalle
        dt2.Clear()

        'Agrega a la tabla de detalle, la tabla que se recibe como parametro
        For Each dr In dt.Rows
            dr2 = dt2.NewRow
            For i = 1 To dt.Columns.Count - 1
                Select Case i
                    Case 1  'Campo NUM_0 (número de intervención)
                        dr2(i) = Numero

                    Case 8  'Campo FACTURA_0 (Menu Local 1)
                        dr2(i) = CInt(dr(i)) + 1

                    Case 11
                        dr2(i) = Me.SolicitudAsociada.Numero

                    Case Else
                        dr2(i) = dr(i)

                End Select

            Next
            dt2.Rows.Add(dr2)
        Next

    End Sub
    Public Sub AgregarDetalle(ByVal Articulo As String, ByVal Cantidad As Integer, ByVal Factura As Boolean, ByVal Tipo As Integer, ByVal Precio As Double)
        Dim dr As DataRow

        dr = dt2.NewRow
        dr("num_0") = Numero
        dr("itmref_0") = Articulo
        dr("qty_0") = 0
        dr("tqty_0") = Cantidad
        dr("factura_0") = IIf(Factura, 2, 1)
        dr("typlig_0") = Tipo
        dr("amt_0") = Precio
        dr("srenum_0") = Me.SolicitudAsociada.Numero
        dt2.Rows.Add(dr)

    End Sub
    Public Sub MarcarEquipos(ByVal FechaInicio As Date, Optional ByVal FechaFin As Date = #12/31/1599#)
        Dim Articulos As New ArrayList

        For Each dr As DataRow In dt2.Rows
            If CInt(dr("typlig_0")) = 1 Then
                Articulos.Add(dr("itmref_0").ToString)
            End If
        Next

        MarcarEquipos(Articulos, FechaInicio, FechaFin)

    End Sub
    Public Sub MarcarEquipos(ByVal Articulos As ArrayList, ByVal FechaInicio As Date, Optional ByVal FechaFin As Date = #12/31/1599#)
        'Esta función marca el campo XITN_0 de los matafuegos con el número de intervención
        'con los que se van a procesar
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim dr As DataRow

        If FechaFin = #12/31/1599# Then FechaFin = FechaInicio.AddMonths(1)

        'daParque
        Sql = "SELECT mac.* "
        Sql &= "FROM (((machines mac INNER JOIN ymacitm ymc ON (mac.macnum_0 = ymc.macnum_0)) INNER JOIN bomd bmd ON (macpdtcod_0 = itmref_0 AND ymc.cpnitmref_0 = bmd.cpnitmref_0)) INNER JOIN bpcustomer bpc ON (mac.bpcnum_0 = bpc.bpcnum_0)) INNER JOIN itmmaster itm ON (ymc.cpnitmref_0 = itm.itmref_0) "
        Sql &= "WHERE bomalt_0 = 99 AND "
        Sql &= "      bomseq_0 = 10 AND "
        Sql &= "      macitntyp_0 = 1 AND "
        Sql &= "      datnext_0 >= to_date(:datnext_0, 'dd/mm/yyyy') AND "
        Sql &= "      datnext_0 <  to_date(:datnext_1, 'dd/mm/yyyy') AND "
        Sql &= "      mac.bpcnum_0 = :bpcnum_0 AND "
        Sql &= "      mac.fcyitn_0 = :fcyitn_0 AND "
        Sql &= "      mac.xitn_0 IN (' ', 'X0') AND "
        Sql &= "      cpnitmref_0 = :cpnitmref_0"

        da = New OracleDataAdapter(Sql, cn)
        With da.SelectCommand.Parameters
            .Add("datnext_0", OracleType.VarChar) 'Fecha Desde
            .Add("datnext_1", OracleType.VarChar) 'Fecha Hasta 30 dias
            .Add("bpcnum_0", OracleType.VarChar)  'Cliente
            .Add("fcyitn_0", OracleType.VarChar)  'Sucursal
            .Add("cpnitmref_0", OracleType.VarChar) 'Componente
        End With
        Sql = "UPDATE machines SET xitn_0 = :xitn_0 WHERE macnum_0 = :macnum_0"
        da.UpdateCommand = New OracleCommand(Sql, cn)
        With da
            Parametro(.UpdateCommand, "xitn_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "macnum_0", OracleType.VarChar, DataRowVersion.Original)
        End With

        'Recupero los equipos a marcar del detalle de la intervencion
        For Each Articulo As String In Articulos
            With da.SelectCommand
                .Parameters("datnext_0").Value = FechaInicio.ToShortDateString
                .Parameters("datnext_1").Value = FechaFin.ToShortDateString
                .Parameters("bpcnum_0").Value = dt1.Rows(0).Item("bpc_0").ToString
                .Parameters("fcyitn_0").Value = dt1.Rows(0).Item("bpaadd_0").ToString
                .Parameters("cpnitmref_0").Value = Articulo
            End With

            da.Fill(dt)
        Next

        'Marco los equipos con el numero de intervencion
        For Each dr In dt.Rows
            dr.BeginEdit()
            dr("xitn_0") = Numero
            dr.EndEdit()
        Next

        da.Update(dt)

        dt.Dispose()
        da.Dispose()

    End Sub
    Public Sub EnvioMailAvisoCordinacion()
        Dim eMail As New CorreoElectronico
        Dim sih As New Factura(cn)
        Dim rpt As New ReportDocument

        ' Salgo si no encuentro la factura
        If Not sih.AbrirPorSolicitud(Me.SolicitudAsociada.Numero) Then
            Exit Sub
        End If

        With rpt
            .Load(RPTX3 & "XFACT_ELEC.rpt") 'Reporte normal
            .SetDatabaseLogon(DB_USR, DB_PWD)
            .SetParameterValue("facturedeb", sih.Numero)
            .SetParameterValue("facturefin", sih.Numero)
            .SetParameterValue("CERO", False)
            .SetParameterValue("OCULTAR_BARRAS", True)
            .SetParameterValue("ENVIAR", True)
            .ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, sih.Numero & ".pdf")
        End With

        With eMail
            .Remitente("contados@georgia.com.ar", "Georgia Seguridad contra Incendios")
            .AgregarDestinatario("contados@georgia.com.ar", True)
            .AgregarDestinatario(Me.Cliente.MailFC)

            .Asunto = "Entrega de Recargas"
            .EsHtml = True
            .CuerpoDesdeArchivo("plantillas\entrega-de-recargas.html")
            .Cuerpo = .Cuerpo.Replace("{cliente}", Me.Tercero.Nombre)
            .Cuerpo = .Cuerpo.Replace("{importe}", sih.ImporteII.ToString("N2"))
            .AdjuntarArchivo(sih.Numero & ".pdf")

            If .CantidadTo > 0 Then .Enviar()

            .Dispose()
        End With

        Try
            File.Delete(sih.Numero & ".pdf")

        Catch ex As Exception
        End Try

    End Sub
    Public Sub EnvioMailAvisoMostrador()
        Dim txt As String = ""
        Dim eMail As New CorreoElectronico
        Dim sih As New Factura(cn)
        Dim rep As New Vendedor(cn)
        Dim sr As New StreamReader("plantillas\aviso-de-recargas-listas.html")

        'Salgo si no encuentro la factura
        If Not sih.AbrirPorSolicitud(Me.SolicitudAsociada.Numero) Then
            Exit Sub
        End If

        Try
            'Verifico que la factura tenga CAE
            If sih.EsFacturaElectronica AndAlso sih.CAE = "" Then
                txt = "Se transfirio la intervención " & Me.Numero
                txt &= vbCrLf & vbCrLf
                txt &= "La factura asociada " & sih.Numero & " NO tiene CAE"
                txt &= vbCrLf & vbCrLf
                txt &= "Mail enviado al clilente, tiene factura adjunta en blanco"

                eMail.Nuevo()
                eMail.Remitente("info@georgia.com.ar")
                eMail.AgregarDestinatarioArchivo("MAILS\sin-cae.txt")
                eMail.Cuerpo = txt
                eMail.EsHtml = False
                eMail.Enviar()
                Exit Sub
            End If
        Catch ex As Exception
        End Try


        txt = sr.ReadToEnd
        txt = txt.Replace("{cliente}", Me.Tercero.Nombre)
        txt = txt.Replace("{fecha}", Me.FechaCreacion.ToString("dd/MM/yyyy"))
        txt = txt.Replace("{itn}", Me.Numero)
        txt = txt.Replace("{importe}", sih.ImporteII.ToString("N2"))

        If ExisteConsumo("601003") Then
            'Esta linea se agrega si la intervencion contiene consumo de sustituto
            txt = txt.Replace("<!--PRESTAMOS-->", "Por favor, no olvide traer con usted los matafuegos de préstamo que le entregamos cuando nos dejó sus equipos.")
        End If

        rep = CType(Me.Cliente, Cliente).Vendedor 'Obtengo vendedor

        If rep.Codigo = "17" Then
            txt = txt.Replace("{vendedor}", rep.Analista.Nombre.ToUpper)
            txt = txt.Replace("{interno}", "")
            txt = txt.Replace("{mail}", rep.Mail)
        Else
            txt = txt.Replace("{vendedor}", rep.Nombre.ToUpper)
            txt = txt.Replace("{interno}", rep.Interno)
            txt = txt.Replace("{mail}", rep.Mail)
        End If

        Dim rpt As New ReportDocument
        With rpt
            .Load(RPTX3 & "XFACT_ELEC.rpt") 'Reporte normal
            .SetDatabaseLogon(DB_USR, DB_PWD)
            .SetParameterValue("facturedeb", sih.Numero)
            .SetParameterValue("facturefin", sih.Numero)
            .SetParameterValue("CERO", False)
            .SetParameterValue("OCULTAR_BARRAS", True)
            .SetParameterValue("ENVIAR", True)
            .ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, sih.Numero & ".pdf")
        End With
        Try
            With eMail
                .Remitente(rep.Mail, rep.Nombre)
                .AgregarDestinatario(Me.Cliente.MailFC)
                .AgregarDestinatarioCopia(rep.Mail)
                .Asunto = "Aviso de recargas listas"
                .EsHtml = True
                .Cuerpo = txt
                .AdjuntarArchivo(sih.Numero & ".pdf")
                If .CantidadTo > 0 Then .Enviar(True)
            End With

        Catch ex As Exception

        End Try


    End Sub
    Public Sub EnvioMailAviso(ByVal ruta As Integer, ByVal motivo As String, ByVal obs As String)
        Dim txt As String = ""
        Dim eMail As New CorreoElectronico
        Dim sih As New Factura(cn)
        Dim rep As New Vendedor(cn)
        rep = Me.Cliente.Vendedor 'Obtengo vendedor

        txt &= "<p> En la siguiente intervencion: {intervencion}, del tipo: {tipo}, tuvimos un rebote en la ruta</p>" & vbCrLf
        txt &= "<p>El cliente es {cliente}, de la sucursal: {sucursal}</p>" & vbCrLf
        txt &= "<p>La ruta es: {ruta}, motivo: {motivo} y las observaciones: {obs}</p>"

        'Reemplazo de marcas
        txt = txt.Replace("{intervencion}", Me.Numero)
        txt = txt.Replace("{tipo}", Me.Tipo)
        ' txt = txt.Replace("{fecha}", Me.FechaCreacion.ToString("dd/MM/yyyy"))
        txt = txt.Replace("{cliente}", Me.Cliente.Codigo & " - " & Me.Cliente.Nombre)
        txt = txt.Replace("{sucursal}", Me.Sucursal.Sucursal & " - " & Me.Sucursal.Direccion)
        txt = txt.Replace("{ruta}", ruta.ToString)
        txt = txt.Replace("{motivo}", motivo)
        txt = txt.Replace("{obs}", obs)

        Try
            With eMail
                .Remitente("no-responder@georgia.com.ar", "Georgia Seguridad contra Incendios")
                .AgregarDestinatario(rep.Mail.ToString)
                .AgregarDestinatario(rep.Analista.Mail.ToString)
                .AgregarDestinatario("glorenzo@georgia.com.ar")
                .Asunto = "Intervencion a Resolver"
                .EsHtml = True
                .Cuerpo = txt
                If .CantidadTo > 0 Then .Enviar(True)
            End With

        Catch ex As Exception

        End Try


    End Sub
    Public Sub AgregarConsumo(ByVal itm As Articulo, ByVal cantidad As Integer, ByVal fechaplan As Date, ByVal importe As Double)

        Dim dr As DataRow

        dr = dt5.NewRow
        dr("HDTNUM_0") = MaxHDT()
        dr("SRENUM_0") = Me.SolicitudAsociada.Numero
        dr("ITNNUM_0") = Numero
        dr("TPL_0") = 1
        dr("HDTMACSRE_0") = " "
        dr("HDTMACSET_0") = " "
        dr("HDTCPN_0") = " "
        dr("HDTTYP_0") = 3
        dr("HDTITM_0") = itm.Codigo
        dr("HDTQTY_0") = cantidad
        dr("HDTUOM_0") = itm.UnidadVta
        dr("HDTSTOFCY_0") = " "
        dr("HDTAVA_0") = 0
        dr("HDTAVAUOM_0") = " "
        dr("HDTAUSTYP_0") = 1
        dr("HDTAUS_0") = 8
        dr("HDTPLNDAT_0") = fechaplan
        dr("HDTDONDAT_0") = Date.Today.Date
        dr("HDTDONHOU_0") = 53011
        dr("SPGTIMHOU_0") = cantidad
        dr("SPGTIMMNT_0") = 0
        dr("HDTTYPRUU_0") = 0
        dr("HDTSTOISS_0") = 1
        dr("HDTISSQTY_0") = 0
        dr("HDTISSISS_0") = 0
        dr("YDONFLG_0") = 2
        dr("HDTINV_0") = 2
        dr("HDTTEX_0") = " "
        dr("HDTSALTEX_0") = " "
        dr("HDTAMT_0") = importe
        dr("HDTAMTINV_0") = (importe * cantidad)
        dr("HDTCUR_0") = "ARS"
        dr("INVPITFLG_0") = 0
        dr("VAT_0") = " "
        dr("VAT_1") = " "
        dr("VAT_2") = " "
        dr("CLCAMT1_0") = 0
        dr("CLCAMT2_0") = 0
        dr("DISCRGVAL1_0") = 0
        dr("DISCRGVAL2_0") = 0
        dr("DISCRGVAL3_0") = 0
        dr("DISCRGVAL4_0") = 0
        dr("DISCRGVAL5_0") = 0
        dr("DISCRGVAL6_0") = 0
        dr("DISCRGVAL7_0") = 0
        dr("DISCRGVAL8_0") = 0
        dr("DISCRGVAL9_0") = 0
        dr("DISCRGREN1_0") = 0
        dr("DISCRGREN2_0") = 0
        dr("DISCRGREN3_0") = 0
        dr("DISCRGREN4_0") = 0
        dr("DISCRGREN5_0") = 0
        dr("DISCRGREN6_0") = 0
        dr("DISCRGREN7_0") = 0
        dr("DISCRGREN8_0") = 0
        dr("DISCRGREN9_0") = 0
        dr("CCE1_0") = " "
        dr("CCE2_0") = "SERVICE"
        dr("CCE3_0") = " "
        dr("CCE4_0") = " "
        dr("CCE5_0") = " "
        dr("CCE6_0") = " "
        dr("CCE7_0") = " "
        dr("CCE8_0") = " "
        dr("CCE9_0") = " "
        dr("PRIREN_0") = 2
        dr("MANAMTFLG_0") = 1
        dr("SAUSTUCOE_0") = 1
        dr("HDTSTUQTY_0") = cantidad
        dr("HDTINVMOD_0") = 0
        dr("YHDITMGEN_0") = " "
        dr("YTRIPNUM_0") = " "
        dr("YHDTAMTGRO_0") = importe
        dt5.Rows.Add(dr)

        dr = dt6.NewRow
        dr("SRENUM_0") = Me.SolicitudAsociada.Numero
        dr("HDTORD_0") = 0
        dr("HDTITM_0") = itm.Codigo
        dr("HDTQTY_0") = cantidad
        dr("HDTUOM_0") = itm.UnidadVta
        dr("HDTTEX_0") = " "
        dr("HDTSALTEX_0") = " "
        dr("HDTAMTINV_0") = importe
        dr("HDTCUR_0") = "ARS"
        dr("VAT_0") = " "
        dr("VAT_1") = " "
        dr("VAT_2") = " "
        dr("CLCAMT1_0") = 0
        dr("CLCAMT2_0") = 0
        dr("DISCRGVAL1_0") = 0
        dr("DISCRGVAL2_0") = 0
        dr("DISCRGVAL3_0") = 0
        dr("DISCRGVAL4_0") = 0
        dr("DISCRGVAL5_0") = 0
        dr("DISCRGVAL6_0") = 0
        dr("DISCRGVAL7_0") = 0
        dr("DISCRGVAL8_0") = 0
        dr("DISCRGVAL9_0") = 0
        dr("DISCRGREN1_0") = 0
        dr("DISCRGREN2_0") = 0
        dr("DISCRGREN3_0") = 0
        dr("DISCRGREN4_0") = 0
        dr("DISCRGREN5_0") = 0
        dr("DISCRGREN6_0") = 0
        dr("DISCRGREN7_0") = 0
        dr("DISCRGREN8_0") = 0
        dr("DISCRGREN9_0") = 0
        dr("PRIREN_0") = 2
        dr("SAUSTUCOE_0") = 1
        dr("HDTSTUQTY_0") = cantidad
        dr("CCE1_0") = " "
        dr("CCE2_0") = "SERVICE"
        dr("CCE3_0") = " "
        dr("CCE4_0") = " "
        dr("CCE5_0") = " "
        dr("CCE6_0") = " "
        dr("CCE7_0") = " "
        dr("CCE8_0") = " "
        dr("CCE9_0") = " "
        dt6.Rows.Add(dr)

    End Sub

    'FUNCTION
    Public Function Abrir(ByVal id As String) As Boolean Implements IRuteable.Abrir
        dt1.Clear()
        dt2.Clear()
        dt3.Clear()
        dt4.Clear()
        dt5.Clear()
        dt6.Clear()
        dt7 = Nothing

        l_Equipos = 0
        l_Mangas = 0
        l_Peso = 0
        l_Peso2 = 0 'Peso para unigis sin prestamos
        l_EsTarea = False
        l_PrestamosExt = 0
        l_PrestamosMan = 0
        l_RechazosExt = 0
        l_RechazosMan = 0
        l_TieneCarro = False
        l_Varios = False

        'Abro una intervencion
        da1.SelectCommand.Parameters("num_0").Value = id.ToUpper
        da1.Fill(dt1)

        'Abro el detalle
        da2.SelectCommand.Parameters("num_0").Value = id.ToUpper
        da2.Fill(dt2)

        'Abro las observaciones de la intervencion
        da3.SelectCommand.Parameters("num_0").Value = id.ToUpper
        da3.Fill(dt3)

        'Abro para agregar los comentarios en el fichero
        da4.SelectCommand.Parameters("ident1_0").Value = id.ToUpper
        da4.Fill(dt4)

        'Abro tabla de consumos
        da5.SelectCommand.Parameters("itnnum_0").Value = id.ToUpper
        da5.Fill(dt5)

        If dt1.Rows.Count = 1 Then
            CalcularIdentidad()
            AnalizarIntervencion()
            RaiseEvent IntervencionAbierta(Me)
        End If

        l_Cliente = Nothing
        l_Sucursal = Nothing

        Return (dt1.Rows.Count = 1)

    End Function
    Public Function AbrirRemito(ByVal Numero As String) As Boolean
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim Sql As String

        Sql = "SELECT num_0 FROM interven WHERE ysdhdeb_0 = :ysdhdeb"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("ysdhdeb", OracleType.VarChar).Value = Numero
        da.Fill(dt)
        da.Dispose()

        l_Cliente = Nothing
        l_Sucursal = Nothing

        If dt.Rows.Count = 0 Then
            Return False
        Else
            Dim dr As DataRow = dt.Rows(0)
            Return Abrir(dr("num_0").ToString)
        End If
    End Function
    Private Function NuevoNumero(ByVal Planta As String, Optional ByVal ModoTest As Boolean = False) As String
        Dim Serie As Integer
        Dim dr As OracleDataReader
        Dim flg As Boolean = (cn.State = ConnectionState.Closed)

        Dim cm1 As New OracleCommand("SELECT valeur_0 FROM avalnum where codnum_0 = 'ACT'", cn)
        Dim cm2 As New OracleCommand("UPDATE avalnum SET valeur_0 = :valeur_0 WHERE codnum_0 = 'ACT'", cn)
        cm2.Parameters.Add("valeur_0", OracleType.Number)

        dr = cm1.ExecuteReader(CommandBehavior.SingleResult)
        dr.Read()
        Serie = CType(dr(0), Integer)
        dr.Close()

        'Aumento numerador
        cm2.Parameters("valeur_0").Value = Serie + 1

        'Si es test no adelanta el numerador en la tabla de contadores
        If Not ModoTest Then cm2.ExecuteNonQuery()

        cm1.Dispose()
        cm2.Dispose()
        dr.Dispose()

        Dim s As String = String.Format("00000000{0}", Serie)

        Return Planta & s.Substring(s.Length - 8)

    End Function
    Friend Function Nueva(ByVal ss As Solicitud, ByVal Tipo As String, ByVal sucursal As String, Optional ByVal diadesde As String = "", Optional ByVal diahasta As String = "", Optional ByVal tardedesde As String = "", Optional ByVal tardehasta As String = "") As Intervencion
        Dim Suc As New Sucursal(cn, ss.Tercero.Codigo, sucursal)
        Dim dr As DataRow

        dt1.Clear()
        dt2.Clear()
        dt3.Clear()

        l_Cliente = Nothing
        l_Sucursal = Nothing
        l_Varios = False

        'Valido si la sucursal es de entrega
        If Suc.SucursalEntregaActiva Then

            dr = dt1.NewRow

            dr("num_0") = NuevoNumero(ss.PlantaAlmacenamiento, False)
            dr("salfcy_0") = ss.PlantaAlmacenamiento
            dr("srvdemnum_0") = ss.Numero
            dr("bpc_0") = ss.Tercero.Codigo
            dr("ccn_0") = ss.Contacto
            dr("mac_0") = " "
            dr("macgru_0") = " "
            dr("typ_0") = Tipo
            dr("dat_0") = Date.Today.AddDays(1)
            dr("datend_0") = Date.Today.AddDays(1)
            dr("hou_0") = "0900"
            dr("houend_0") = "1000"
            dr("fulday_0") = 1
            dr("wee_0") = 1
            dr("dur_0") = "0100"
            dr("datx_0") = Date.Today.AddDays(1)
            dr("dtckil_0") = 0
            dr("tritim_0") = 0
            dr("obj_0") = " "
            dr("typfulobj_0") = "ITNOBJ"
            dr("numfulobj_0") = dr("num_0", DataRowVersion.Proposed).ToString
            dr("objflg_0") = 0
            dr("rep_0") = USER
            dr("timspg_0") = 0
            dr("htottimspg_0") = 0
            dr("mtottimspg_0") = 0
            dr("mantimflg_0") = 0
            dr("pblsol_0") = 1
            dr("srvconcov_0") = 0
            dr("connum_0") = " "
            dr("ordnum_0") = " "
            dr("sco_0") = 1
            dr("sconum_0") = " "
            dr("scoamt_0") = 0
            dr("cur_0") = " "
            dr("bpaadd_0") = Suc.Sucursal
            dr("drn_0") = Suc.Ruta
            dr("itntypadd_0") = 99
            dr("itncodadd_0") = " "
            dr("itnrecadd_0") = " "
            dr("add_0") = Suc.Direccion
            dr("add_1") = " "
            dr("add_2") = " "
            dr("zip_0") = IIf(Suc.CodigoPostal = "", " ", Suc.CodigoPostal)
            dr("cty_0") = Suc.Ciudad
            dr("cry_0") = Suc.Pais.Codigo
            dr("sat_0") = Suc.Provincia
            dr("tel_0") = Suc.Telefono
            dr("xportero_0") = Suc.Portero
            dr("xpor_tel_0") = Suc.Telefono_Portero
            dr("xmailfc_0") = Suc.MailFC
            dr("iffadd_0") = " "
            dr("rer_0") = " "
            dr("rer_1") = " "
            dr("rer_2") = " "
            dr("rer_3") = " "
            dr("rer_4") = " "
            dr("rer_5") = " "
            dr("rer_6") = " "
            dr("rer_7") = " "
            dr("rer_8") = " "
            dr("rer_9") = " "
            dr("rer_10") = " "
            dr("rer_11") = " "
            dr("rer_12") = " "
            dr("rer_13") = " "
            dr("rer_14") = " "
            dr("don_0") = 1
            dr("typfulrpo_0") = "ITN"
            dr("numfulrpo_0") = dr("num_0", DataRowVersion.Proposed).ToString
            dr("rpo_0") = " "
            dr("rpoflg_0") = 0
            dr("itnori_0") = 1
            dr("itnoritxt_0") = "Creacion manual"
            dr("itnorivcr_0") = " "
            dr("itnorivcrl_0") = 0
            dr("creusr_0") = USER
            dr("credat_0") = Date.Today
            dr("crehou_0") = Date.Now.ToString("HHmm")
            dr("updusr_0") = " "
            dr("upddat_0") = #12/31/1599#
            dr("ypctpres_0") = 0
            dr("ycllnum_0") = " "
            dr("zflgtrip_0") = 1
            dr("yflgsdh_0") = 0
            dr("yhdtamtinv_0") = 0
            dr("ymrkitn_0") = 2
            dr("tripnum_0") = " "
            dr("ysdhdeb_0") = " "
            dr("ysdhfin_0") = " "
            dr("mdl_0") = Suc.ModoEntrega
            dr("yref_0") = " "
            dr("dlvpio_0") = Suc.PrioridadEntrega
            dr("yflgret_0") = 1
            dr("yhdesde1_0") = IIf(diadesde = "", Suc.TurnoMananaDesde, diadesde)
            dr("yhhasta1_0") = IIf(diahasta = "", Suc.TurnoMananaHasta, diahasta)
            dr("yhdesde2_0") = IIf(tardedesde = "", Suc.TurnoTardeDesde, tardedesde)
            dr("yhhasta2_0") = IIf(tardehasta = "", Suc.TurnoTardeHasta, tardehasta)
            dr("yotr_0") = 0
            dr("xauto_0") = 1
            dr("ymesa_0") = " "
            dr("yobserva_0") = " "
            dr("yobsrec_0") = " "
            dr("yobsitn_0") = " "
            dr("xnoconf_0") = " "
            dr("xsector_0") = " "
            dr("xweb_0") = " "
            dr("xtrans_0") = 1
            dr("xgeo_0") = 2 'indica que la intervención se creo desde georgia
            dr("xrpt_0") = 1 'indica reporte ctrl. periodico no enviado aún.
            dr("xcer_0") = 0
            dr("xtanda_0") = 1
            dr("ocf_0") = " "
            dr("xcarzon_0") = 0
            dr("xcardat_0") = #12/31/1599#
            dr("xcartyp_0") = 0
            dt1.Rows.Add(dr)

            Return Me

        Else
            Return Nothing

        End If

    End Function
    Public Function ExisteRetira(ByVal Articulo As String) As Boolean
        'Devuelve Verdadero o Falso si el codigo de Articulo figura en la seccion retira de la intervencion
        Dim dv As New DataView(dt2)
        dv.RowFilter = "itmref_0 = '" & Articulo & "'"

        Return (dv.Count <> 0)
    End Function
    Public Function ExisteConsumo(ByVal Articulo As String) As Boolean
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable

        Sql = "SELECT * "
        Sql &= "FROM hdktask "
        Sql &= "WHERE srenum_0 = :srenum_0 AND itnnum_0 = :itnnum_0 AND hdtitm_0 = :hdtitm_0"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("srenum_0", OracleType.VarChar).Value = SolicitudAsociada.Numero
        da.SelectCommand.Parameters.Add("itnnum_0", OracleType.VarChar).Value = Numero
        da.SelectCommand.Parameters.Add("hdtitm_0", OracleType.VarChar).Value = Articulo
        da.Fill(dt)
        da.Dispose()

        Return dt.Rows.Count > 0

    End Function
    Public Function RetiroSaldo() As DataTable
        'Devuelve una tabla con la estructua de YITNDET con las cantidades faltantes de retirar
        Dim dt As New DataTable

        da2.FillSchema(dt, SchemaType.Mapped)

        'Clono el objeto table
        For Each dr2 As DataRow In dt2.Rows
            Dim dr As DataRow

            dr = dt.NewRow
            For i = 0 To dt2.Columns.Count - 1
                dr(i) = dr2(i)
            Next
            dt.Rows.Add(dr)
        Next

        Dim tot As Integer = 0

        'Por cada registro consulto la cantidad ingresada por Godoy
        For Each dr As DataRow In dt.Rows
            If CInt(dr("typlig_0")) <> 1 Then Continue For

            Dim c1 As Integer
            Dim c2 As Integer
            Dim c3 As Integer

            c1 = IngresadoPorGodoy(dr("itmref_0").ToString)
            c2 = CInt(dr("tqty_0"))
            c3 = c2 - c1

            dr.BeginEdit()
            dr("tqty_0") = IIf(c3 <= 0, 0, c3)
            dr.EndEdit()

            tot += CInt(IIf(c3 <= 0, 0, c3))
        Next

        'Elimino los registros con cantidades < 0
        For i As Integer = dt.Rows.Count - 1 To 0 Step -1
            Dim dr1 As DataRow = dt.Rows(i)
            If CInt(dr1("tqty_0")) <= 0 Then dt.Rows.Remove(dr1)
        Next

        'Prestamos
        For Each dr As DataRow In dt.Rows
            If CInt(dr("typlig_0")) = 2 Then
                dr.BeginEdit()
                dr("tqty_0") = tot
                dr.EndEdit()
            End If

        Next

        Return dt

    End Function
    Public Function TieneRechazos() As Boolean
        Dim dt As New DataTable
        Dim da As OracleDataAdapter
        Dim Sql As String = ""

        Sql &= "select mac.macnum_0 "
        Sql &= "from sremac sre inner join "
        Sql &= "	 machines mac on (sre.macnum_0 = mac.macnum_0) "
        Sql &= "where mac.macitntyp_0 = 5 and "
        Sql &= "	  yitnnum_0 = :itn "
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("itn", OracleType.VarChar).Value = Me.Numero
        da.Fill(dt)
        da.Dispose()

        Return dt.Rows.Count > 0

    End Function
    Public Sub ActualizarRetiroTeorico()
        'Sobreescribe los valores de retiros teoricos con los valores reales ingresados por Godoy
        For Each dr As DataRow In dt2.Rows
            If CInt(dr("typlig_0")) <> 1 Then Continue For
            dr.BeginEdit()
            dr("tqty_0") = IngresadoPorGodoy(dr("itmref_0").ToString)
            dr.EndEdit()
        Next
        da2.Update(dt2)

    End Sub
    Private Function IngresadoPorGodoy(ByVal Codigo As String) As Integer
        'Devuelve la cantidad de extintores ingresados por Godoy que tengan el codigo de vencimiento consultado
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim sql As String
        Dim cant As Integer = 0

        sql = "SELECT COUNT(ymc.cpnitmref_0) AS canti "
        sql &= "FROM machines mac INNER JOIN "
        sql &= "     ymacitm ymc ON (mac.macnum_0 = ymc.macnum_0) INNER JOIN "
        sql &= "     bomd bmd ON (mac.macpdtcod_0 = bmd.itmref_0 AND ymc.cpnitmref_0 = bmd.cpnitmref_0 AND bomalt_0 = 99 AND bomseq_0 = 10) INNER JOIN "
        sql &= "     sremac sre ON (mac.macnum_0 = sre.macnum_0) "
        sql &= "WHERE sre.yitnnum_0 = :yitnnum AND "
        sql &= "      ymc.cpnitmref_0 = :cpnitmref "
        sql &= "GROUP BY ymc.cpnitmref_0"
        da = New OracleDataAdapter(sql, cn)
        da.SelectCommand.Parameters.Add("yitnnum", OracleType.VarChar).Value = Numero
        da.SelectCommand.Parameters.Add("cpnitmref", OracleType.VarChar).Value = Codigo

        Try
            da.Fill(dt)

        Catch ex As Exception

        End Try

        If dt.Rows.Count = 1 Then
            dr = dt.Rows(0)
            cant = CInt(dr(0))
        End If

        da.Dispose()
        dt.Dispose()

        Return cant

    End Function
    Private Function MaxHDT() As String
        'Variables con el numerador de consumos
        Dim c As New Numerador(cn, "HDT", Me.Planta.SociedadPlanta.Codigo, CInt(Today.ToString("yy")))
        Dim v As Long
        Dim s As String

        'Obtengo el proximo numero de consumo
        v = c.Valor

        'doy formado al numero de consumo
        s = "C"
        s &= Me.Planta.CodigoPlanta
        s &= Today.ToString("yy")
        s &= v.ToString("000000")

        Return s

    End Function

    Shared Function RetirosSchema(ByVal cn As OracleConnection) As DataTable
        Dim dt As New DataTable
        Dim da As OracleDataAdapter
        Dim sql As String

        sql = "SELECT * FROM yitndet"
        da = New OracleDataAdapter(sql, cn)
        Try
            da.FillSchema(dt, SchemaType.Mapped)
            da.Dispose()
        Catch ex As Exception
        Finally

        End Try

        Return dt

    End Function

    'EVENTS
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

        If e.StatementType = StatementType.Update Then
            Dim dr As DataRow = dt1.Rows(0)

            dr.BeginEdit()
            dr("upddat_0") = Date.Today
            dr("updusr_0") = USER
            dr.EndEdit()
        End If
    End Sub
    Private Sub da3_RowUpdated(ByVal sender As Object, ByVal e As System.Data.OracleClient.OracleRowUpdatedEventArgs) Handles da3.RowUpdated
        If e.Status = UpdateStatus.ErrorsOccurred Then Exit Sub

        If e.StatementType = StatementType.Update Then
            Dim dr As DataRow = dt1.Rows(0)
            dr.BeginEdit()
            dr("upddat_0") = Date.Today
            dr("updusr_0") = USER
            dr.EndEdit()
        End If
    End Sub

    'PROPERTY
    Public ReadOnly Property Numero() As String Implements IRuteable.Numero
        Get
            Dim dr As DataRow = dt1.Rows(0)

            Return dr("num_0").ToString
        End Get
    End Property
    Public ReadOnly Property Remito() As String Implements IRuteable.Remito
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("ysdhdeb_0").ToString
        End Get
    End Property
    Public ReadOnly Property SolicitudAsociada() As Solicitud
        Get
            Dim ss As New Solicitud(cn)
            Dim dr As DataRow = dt1.Rows(0)
            ss.Abrir(dr("srvdemnum_0").ToString)
            Return ss
        End Get
    End Property
    Public ReadOnly Property SolicitudServicio() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("srvdemnum_0").ToString()
        End Get
    End Property
    Public ReadOnly Property Sucursal() As Sucursal Implements IRuteable.Sucursal
        Get
            If l_Sucursal Is Nothing Then
                l_Sucursal = New Sucursal(cn)
                l_Sucursal.Abrir(Me.Cliente.Codigo, Me.SucursalCodigo)
            End If

            Return l_Sucursal

        End Get
    End Property
    Public ReadOnly Property SucursalCodigo() As String Implements IRuteable.SucursalCodigo
        Get
            Dim dr As DataRow = dt1.Rows(0)

            Return dr("bpaadd_0").ToString
        End Get
    End Property
    Public ReadOnly Property SucursalCalle() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("add_0").ToString
        End Get
    End Property
    Public ReadOnly Property SucursalCiudad() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("cty_0").ToString
        End Get
    End Property
    Public ReadOnly Property SucursalProvincia() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("sat_0").ToString
        End Get
    End Property
    Public ReadOnly Property Planta() As Planta
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Dim fcy As New Planta(cn)

            fcy.Abrir(dr("salfcy_0").ToString)

            Return fcy

        End Get
    End Property
    Public ReadOnly Property Tercero() As Tercero Implements IRuteable.Tercero
        Get
            If l_Cliente Is Nothing Then
                Dim dr As DataRow = dt1.Rows(0)

                l_Cliente = New Cliente(cn)
                l_Cliente.Abrir(dr("bpc_0").ToString)
            End If

            Return l_Cliente
        End Get
    End Property
    Public ReadOnly Property Cliente() As Cliente
        Get
            Return CType(Me.Tercero, Cliente)
        End Get
    End Property
    Public Property Tipo() As String Implements IRuteable.Tipo
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("typ_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt1.Rows(0)
            dr.BeginEdit()
            dr("typ_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Observaciones() As String
        Get
            Dim txt As String = ""
            If dt3.Rows.Count = 1 Then txt = dt3.Rows(0).Item("clob_0").ToString
            Return txt
        End Get
        Set(ByVal value As String)
            Dim dr1 As DataRow = dt1.Rows(0)
            Dim dr3 As DataRow

            If value.Trim = "" Then
                dr1.BeginEdit()
                dr1("obj_0") = " "
                dr1.EndEdit()

                If dt3.Rows.Count = 1 Then
                    dt3.Rows(0).Delete()
                End If

            Else

                dr1.BeginEdit()
                If value.Length > 230 Then
                    dr1("obj_0") = value.Substring(0, 230)
                Else
                    dr1("obj_0") = value
                End If

                dr1.EndEdit()

                If dt3.Rows.Count = 0 Then
                    dr3 = dt3.NewRow
                    dr3("num_0") = Me.Numero
                    dr3("typ_0") = "ITNOBJ"
                    dr3("clob_0") = value
                    dt3.Rows.Add(dr3)
                Else
                    dr3 = dt3.Rows(0)

                    If dr3.RowState = DataRowState.Deleted Then dr3.RejectChanges()
                    dr3.BeginEdit()
                    dr3("clob_0") = value
                    dr3.EndEdit()
                End If

            End If
        End Set
    End Property
    Public Property ComentarioRec() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("yobsrec_0").ToString.Trim
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt1.Rows(0)
            dr.BeginEdit()
            If value = "" Then value = " "
            dr("yobsrec_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property ComentarioItn() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("yobsitn_0").ToString.Trim
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt1.Rows(0)
            dr.BeginEdit()
            If value = "" Then value = " "
            dr("yobsitn_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Estado() As Integer
        Get
            Dim dr As DataRow

            dr = dt1.Rows(0)
            Return CInt(dr("zflgtrip_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("zflgtrip_0") = value
            dr.EndEdit()

            'Si la intervencion se cierra se dispara evento de equipos ocultos por itn
            If value = 8 Then
                BuscarEquiposMarcados()
            End If
        End Set
    End Property
    Public Property Efectuado() As Boolean
        Get
            Dim dr As DataRow = dt1.Rows(0)

            If CInt(dr("don_0")) = 2 Then
                Return True
            Else
                Return False
            End If
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow = dt1.Rows(0)

            dr.BeginEdit()
            dr("don_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property SolicitudSatisfecha() As Boolean
        Get
            Dim dr As DataRow = dt1.Rows(0)

            If CInt(dr("pblsol_0")) = 2 Then
                Return True
            Else
                Return False
            End If
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow = dt1.Rows(0)

            dr.BeginEdit()
            If value Then
                dr("don_0") = 2
                dr("pblsol_0") = 2
            Else
                dr("pblsol_0") = 1
            End If
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property FechaUnigis() As Date Implements IRuteable.FechaUnigis
        Get
            Select Case TipoTarea
                Case "ENT", "LOG"
                    'Fecha de pistoleo a Logistica
                    Dim Sectores() As String = {"ABO", "LOG"}
                    Dim p As New Seguimiento(cn)
                    p.Abrir(Me)

                    Return p.UltimaFechaEnviadoA(Sectores)

                Case Else
                    Return FechaInicio

            End Select

        End Get
    End Property
    Public Property FechaInicio() As Date
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return CDate(dr("dat_0"))
        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow = dt1.Rows(0)
            dr.BeginEdit()
            dr("dat_0") = value
            dr.EndEdit()

            If value > FechaFin Then FechaFin = value
        End Set
    End Property
    Public Property FechaFin() As Date
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return CDate(dr("datend_0"))
        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow = dt1.Rows(0)
            dr.BeginEdit()
            dr("datend_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property MotivoNoConformidad() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("xnoconf_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt1.Rows(0)
            dr.BeginEdit()
            dr("xnoconf_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property OTR() As Integer
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CInt(dr("yotr_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("yotr_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property OCF() As String
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return dr("ocf_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("ocf_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Transito() As Boolean
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CBool(CInt(dr("xtrans_0")) = 2)
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("xtrans_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property FechaCreacion() As Date
        Get
            Dim dr As DataRow
            Dim f As Date

            If dt1.Rows.Count = 1 Then
                dr = dt1.Rows(0)
                f = CDate(dr("credat_0"))
            End If

            Return f

        End Get
    End Property
    Public Property Mesa() As String
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return dr("ymesa_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("ymesa_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Ruta() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("tripnum_0").ToString

        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt1.Rows(0)
            dr.BeginEdit()
            dr("tripnum_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Tanda() As Boolean
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return CInt(dr("xtanda_0")) = 2

        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow = dt1.Rows(0)
            dr.BeginEdit()
            dr("xtanda_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Sector() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("xsector_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt1.Rows(0)
            dr.BeginEdit()
            dr("xsector_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property CantidadEquiposTeoricos() As Integer
        Get
            Dim dr As DataRow
            Dim i As Integer = 0

            For Each dr In dt2.Rows
                If CInt(dr("typlig_0")) = 1 Then i += CInt(dr("tqty_0"))
            Next

            Return i

        End Get

    End Property
    Public Property Reclamo() As Boolean
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return CInt(dr("ymrkitn_0")) = 1
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow = dt1.Rows(0)
            dr.BeginEdit()
            dr("ymrkitn_0") = IIf(value, 1, 2)
            dr.EndEdit()
        End Set
    End Property
    Public Property Referencia() As String
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return dr("yref_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("yref_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public ReadOnly Property Detalle() As DataView
        Get
            Dim dv As New DataView(dt2)
            dv.RowFilter = "typlig_0 = 1"
            Return dv
        End Get
    End Property
    Public Property eMail() As String
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return dr("xweb_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("xweb_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property RptEnviado() As Boolean
        Get
            Dim dr As DataRow = dt1.Rows(0)

            If CInt(dr("xrpt_0")) = 2 Then
                Return True
            Else
                Return False
            End If
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow = dt1.Rows(0)

            dr.BeginEdit()
            If value Then
                dr("xrpt_0") = 2
            Else
                dr("xrpt_0") = 1
            End If
            dr.EndEdit()
        End Set
    End Property
    Public Property CertificadoEstado() As Integer
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return CInt(dr("xcer_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow = dt1.Rows(0)
            dr("xcer_0") = value
        End Set
    End Property
    Public Property ficheComment() As String
        Get
            Dim dr As DataRow = dt4.Rows(0)
            Return dr("clob_").ToString
        End Get
        Set(ByVal value As String)
            If dt4.Rows.Count = 0 Then
                Dim dr As DataRow
                dr = dt4.NewRow
                dr("Codblb_0") = "CO_ITN"
                dr("IDENT1_0") = Numero
                dr("IDENT2_0") = " "
                dr("IDENT3_0") = 1
                dr("NAMBLB_0") = "CO_ITN " & Numero
                dr("TYPDOC_0") = "RTF"
                dr("CREUSR_0") = "ADMIN"
                dr("CREDAT_0") = Today.Date
                dr("CRETIM_0") = Today.Hour * 3600 + Today.Minute * 60
                dr("CLOB_0") = TextRTf(value)
                dt4.Rows.Add(dr)
            Else
                Dim dr As DataRow = dt4.Rows(1)
                dr.BeginEdit()
                dr("clob_0") = TextRTf(value)
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property CarritoZona() As Integer
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return CInt(dr("xcarzon_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow = dt1.Rows(0)
            dr.BeginEdit()
            dr("xcarzon_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property CarritoTipo() As Integer
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return CInt(dr("xcartyp_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow = dt1.Rows(0)
            dr.BeginEdit()
            dr("xcartyp_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property CarritoFecha() As Date Implements IRuteable.CarritoFecha
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return CDate(dr("xcardat_0"))
        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow = dt1.Rows(0)
            dr.BeginEdit()
            dr("xcardat_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property Seguimiento() As Seguimiento
        Get
            Dim segto As New Seguimiento(cn)
            segto.Abrir(Me)
            Return segto
        End Get
    End Property
    Public Property Automatico() As Boolean
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return CBool(IIf(CInt(dr("xauto_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow = dt1.Rows(0)
            dr.BeginEdit()
            dr("xauto_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property ModoEntrega() As String Implements IRuteable.ModoEntrega
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return dr("mdl_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("mdl_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Franja1Desde() As String Implements IRuteable.Franja1Desde
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return dr("yhdesde1_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("yhdesde1_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Franja2Desde() As String Implements IRuteable.Franja2Desde
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return dr("yhdesde2_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("yhdesde2_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Franja1Hasta() As String Implements IRuteable.Franja1Hasta
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return dr("yhhasta1_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("yhhasta1_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Franja2Hasta() As String Implements IRuteable.Franja2Hasta
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return dr("yhhasta2_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("yhhasta2_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property FechaEntrega() As Date Implements IRuteable.FechaEntrega
        Get
            Return Me.FechaFin
        End Get
    End Property
    Private Function TextRTf(ByVal txt As String) As String
        Dim texto As String = "{\rtf1\ansi\ansicpg1252\deff0\deflang3082{\fonttbl{\f0\fswiss\fprq2\fcharset0 MS Sans Serif;}}\viewkind4\uc1\pard\f0\fs17" & txt & "\par}}"
        Return texto
    End Function
    Public Function AnalizarEstado() As String
        Dim r1 As New Ruta(cn)
        Dim r2 As New Ruta(cn)
        Dim txt As String

        If ("A1C1".IndexOf(Me.Tipo)) > -1 And Me.Estado = 1 And Me.Ruta.Trim = "" And Me.CarritoFecha = #12/31/1599# Then
            txt = "El pedido de retiro fue realizado al sector correspondiente el día " & Me.FechaInicio.ToString("dd/MM/yyyy") & " y por el momento no tiene ruta asignada. Consultar a Abonos."
            Return txt
        End If

        If ("E1E2F1".IndexOf(Me.Tipo)) > -1 And Me.Estado = 1 And Me.Ruta.Trim = "" And Me.CarritoFecha = #12/31/1599# Then
            txt = "El pedido de retiro fue realizado al sector correspondiente el día " & Me.FechaInicio.ToString("dd/MM/yyyy") & " y por el momento no tiene ruta asignada. Consultar a Sistemas Fijos."
            Return txt
        End If

        If Me.Tipo = "B1" And Me.Estado = 1 And Me.Ruta.Trim = "" And Me.CarritoFecha = #12/31/1599# Then
            txt = "El pedido de retiro fue realizado al sector correspondiente el día " & Me.FechaInicio.ToString("dd/MM/yyyy") & " y por el momento no tiene ruta asignada. Consultar a Logistica."
            Return txt
        End If

        If Me.Tipo = "B1" And Me.Estado = 1 And Me.Ruta.Trim = "" And Me.CarritoFecha <> #12/31/1599# Then
            txt = "Al sector correspondiente se le solicitó retirar el día " & Me.CarritoFecha.ToString("dd/MM/yyyy")
            Return txt
        End If

        If Me.Estado = 1 And Me.Ruta.Trim <> "" Then

            r1.Abrir(CInt(Me.Ruta))

            If r1.Fecha >= Today Then
                txt = "El retiro se realizará el día (fecha ruta), con nro de ruta (nro ruta), creada por (operador)"
                txt = txt.Replace("(fecha ruta)", r1.Fecha.ToString("dd/MM/yyyy"))
                txt = txt.Replace("(nro ruta)", r1.Numero.ToString)
                txt = txt.Replace("(operador)", r1.UsuarioCreacion)
            Else
                txt = "El sector de (operador) indicó que se haría el retiro el día (fecha ruta) con nro ruta (nro ruta), pero aún no indicó el resultado de esta ruta. Consultar al sector acerca del retiro."
                txt = txt.Replace("(fecha ruta)", r1.Fecha.ToString("dd/MM/yyyy"))
                txt = txt.Replace("(nro ruta)", r1.Numero.ToString)
                txt = txt.Replace("(operador)", r1.UsuarioCreacion)
            End If

            Return txt
        End If

        If Me.Tipo = "D1" And Me.Estado = 1 Then
            txt = "Los equipos han sido ingresados el día " & Me.FechaInicio & " con OT " & Me.OTR & " y están pendientes de ser procesados. Se encuentran en sector Service"
            Return txt
        End If

        If Me.Estado = 7 And Me.Remito = " " Then
            If r1.AbrirUltimaRuta(Me.Numero) Then
                txt = "El día " & r1.Fecha.ToString("dd/MM/yyyy") & " se intento realizar el trabajo, pero no se pudo debido a " & r1.MotivoRebote(Me.Numero) & vbCrLf & vbCrLf
                txt &= "OBSERVACIONES: " & r1.Observacion(Me.Numero) & vbCrLf & vbCrLf
                txt &= "Relevar contacto, horario y teléfono para recoordinar."

                Return txt
            End If
        End If

        If Me.Estado = 6 And Me.Remito = " " And Me.Ruta = " " Then

            If r1.AbrirUltimaRuta(Me.Numero) Then
                txt = "El día " & r1.Fecha.ToString("dd/MM/yyyy") & " se intento realizar el trabajo, pero no se pudo debido a " & r1.MotivoRebote(Me.Numero) & vbCrLf & vbCrLf
                txt &= "OBSERVACIONES: " & r1.Observacion(Me.Numero)
                txt &= "Relevar contacto, horario y teléfono para recoordinar."

                Return txt
            End If

        End If

        If Me.Estado = 6 And Me.Remito = " " And Me.Ruta <> " " Then

            r1.Abrir(CInt(Me.Ruta)) 'Abro ruta actual
            r2.AbrirUltimaRuta(Me.Numero, r1.Numero) 'Ultima Ruta

            If r1.Fecha >= Today Then
                txt = "El día " & r2.Fecha.ToString("dd/MM/yyyy") & " se intento realizar el trabajo, pero no se pudo debido a " & r2.MotivoRebote(Me.Numero)
                txt &= "OBSERVACIONES: " & r2.Observacion(Me.Numero)
                txt &= "El día " & r1.Fecha.ToString("dd/MM/yyyy") & " intentaremos pasar nuevamente."

            Else
                txt = "El día " & r2.Fecha.ToString("dd/MM/yyyy") & " se intento realizar el trabajo, pero no se pudo debido a " & r2.MotivoRebote(Me.Numero)
                txt &= "OBSERVACIONES: " & r2.Observacion(Me.Numero)
                txt &= "El día " & r1.Fecha.ToString("dd/MM/yyyy") & " con ruta Nro. " & r1.Numero.ToString & " se indicó que se pasaría nuevamente pero aún no se indicó el resultado de este nuevo intento. Consultar al sector de " & r1.Camioneta.Sector

            End If

            Return txt

        End If

        If Me.Tipo = "F1" And Me.Estado = 2 Then
            If r1.AbrirUltimaRuta(Me.Numero) Then
                txt = "Las mangueras han sido retiradas el día " & r1.Fecha.ToString("dd/MM/yyyy") & " y están pendientes de ser procesadas. Se encuentran en sector Service"
                Return txt
            End If
        End If

        If "B1C1".IndexOf(Me.Tipo) > -1 And Me.Estado = 2 Then
            If r1.AbrirUltimaRuta(Me.Numero) Then
                txt = "Los equipos han sido retirados el día " & r1.Fecha.ToString("dd/MM/yyyy") & " y están pendientes de ser procesados. Se encuentran en sector Service"
                Return txt

            End If

        End If

        If "B1C1D1F1".IndexOf(Me.Tipo) > -1 And Me.Estado = 3 Then
            txt = "Los equipos ya han sido procesados en el área de Service y se encuentran en proceso de generación de remitos y facturación"
            Return txt
        End If

        If "LOG-ABO-ING".IndexOf(Me.Sector) > -1 And Me.Estado = 4 And Me.Ruta = " " Then
            Dim seg As New Seguimiento(cn)
            seg.Abrir(Me)

            txt = "Los equipos se encuentran en cola de espera para ser entregados al cliente desde el día " & seg.UltimaFechaEnviadoA(seg.UltimoSectorDestino).ToString("dd/MM/yyyy")
            txt &= ", pero aún no tienen fecha asignada."

            Return txt
        End If

        If "LOG-ABO-ING".IndexOf(Me.Sector) > -1 And Me.Estado = 4 And Me.Ruta <> " " Then
            r1.Abrir(CInt(Me.Ruta))

            If r1.Fecha >= Today Then
                txt = "La entrega de los equipos se realizará el día " & r1.Fecha.ToString("dd/MM/yyyy") & ", con nro de ruta " & r1.Numero.ToString & ", creada por " & r1.Camioneta.Sector
            Else
                txt = "El sector de " & r1.Camioneta.Sector & " indicó que se haría la entrega el día " & r1.Fecha.ToString("dd/MM/yyyy") & " con nro ruta " & r1.Numero.ToString & ", pero aún no indicó el resultado de esta ruta. Consultar al sector acerca de la entrega"
            End If

            Return txt
        End If

        If Me.Estado = 6 And Me.Remito <> " " And Me.Ruta = " " Then
            If r1.AbrirUltimaRuta(Me.Numero) Then
                txt = "El día " & r1.Fecha.ToString("dd/MM/yyyy") & " se intento realizar la entrega, pero no se pudo debido a " & r1.MotivoRebote(Me.Numero) & vbCrLf & vbCrLf
                txt &= "OBSERVACIONES: " & r1.Observacion(Me.Numero) & vbCrLf & vbCrLf
                txt &= "Relevar contacto, horario y teléfono para recoordinar."
                Return txt
            End If
        End If

        If Me.Estado = 6 And Me.Remito <> " " And Me.Ruta <> " " Then
            r1.Abrir(CInt(Me.Ruta))
            r2.AbrirUltimaRuta(Me.Numero, r1.Numero)

            If r1.Fecha >= Today Then
                txt = "El día " & r2.Fecha.ToString("dd/MM/yyyy") & " se intento realizar la entrega, pero no se pudo debido a " & r2.MotivoRebote(Me.Numero) & vbCrLf & vbCrLf
                txt &= "OBSERVACIONES: " & r2.Observacion(Me.Numero) & vbCrLf & vbCrLf
                txt = "El día " & r1.Fecha.ToString("dd/MM/yyyy") & " intentaremos pasar nuevamente."
            Else
                txt = "El día " & r2.Fecha.ToString("dd/MM/yyyy") & " se intento realizar la entrega, pero no se pudo debido a " & r2.MotivoRebote(Me.Numero) & vbCrLf & vbCrLf
                txt &= "OBSERVACIONES: " & r2.Observacion(Me.Numero) & vbCrLf & vbCrLf
                txt &= "El día " & r1.Fecha.ToString("dd/MM/yyyy") & " con ruta Nro. " & r1.Numero.ToString & " se indicó que se pasaría nuevamente pero aún no se indicó el resultado de este nuevo intento. "
                txt &= "Consultar al sector de " & r1.Camioneta.Sector
            End If

            Return txt

        End If

        If Me.Estado = 7 And Me.Remito <> " " Then
            If r1.AbrirUltimaRuta(Me.Numero) Then
                txt = "El día " & r1.Fecha.ToString("dd/MM/yyyy") & " se intento realizar el trabajo, pero no se pudo debido a " & r1.MotivoRebote(Me.Numero) & vbCrLf & vbCrLf
                txt &= "OBSERVACIONES: " & r1.Observacion(Me.Numero) & vbCrLf & vbCrLf
                txt &= "El documento se encuentra en el sector " & Me.Sector & ". Relevar contacto, horario y teléfono para recoordinar. "

                Return txt
            End If
        End If

        If Me.Estado = 5 Then
            If r1.AbrirUltimaRuta(Me.Numero) Then
                txt = "La tarea solicitada ha sido concluida el día " & r1.Fecha.ToString("dd/MM/yyyy") & " con nro ruta " & r1.Numero & vbCrLf
                txt &= "El documento físicamente se encuentra en " & Me.Sector
                Return txt
            End If

        End If

        If Me.Estado = 8 Then
            txt = "La tarea fue dada de baja" & vbCrLf & vbCrLf
            txt &= Me.Observaciones
            Return txt
        End If

        txt = "El estado de esta intervención no se corresponde con ningun caso real, consultar a Supervisor"
        Return txt

    End Function
    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If disposedValue Then Exit Sub

        If disposing Then
            ' TODO: Liberar otro estado (objetos administrados).
        End If

        ' TODO: Liberar su propio estado (objetos no administrados).
        ' TODO: Establecer campos grandes como Null.
        da1.Dispose()
        dt1.Dispose()
        da3.Dispose()
        dt3.Dispose()

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
    'CLASE
    '======================================================================================
    Public Class ParquesEvenArgs
        Inherits System.EventArgs

        Public Series As New Specialized.StringCollection

        Public Sub New(ByVal dt As DataTable)
            Dim dr As DataRow

            For Each dr In dt.Rows
                Series.Add(dr("macnum_0").ToString)
            Next
        End Sub

    End Class

    Private Sub AnalizarIntervencion()
        If Remito.Trim = "" Then
            AnalizarIntervencionRetiro()
        Else
            AnalizarIntervencionEntrega()
        End If
    End Sub
    Private Sub AnalizarIntervencionRetiro()
        Dim dr As DataRow
        Dim itm As New Articulo(cn)

        'Intervencion sin remito. Se trata de un retiro de recarga (RET)
        For Each dr In dt2.Rows
            Dim Qty As Integer = CInt(dr("tqty_0"))

            itm.Abrir(dr("itmref_0").ToString)

            If itm.EsCarro Then l_TieneCarro = True
            Dim x As String = Numero

            'Recorro todos los codigos cargados en la solapa Retiro de la intervención
            Select Case CType(dr("typlig_0"), Integer)
                Case 1  'Solapa Retiro

                    Select Case itm.LineaProducto 'dr("cfglin_0").ToString
                        Case "451", "459"  'Recarga equipos
                            l_Equipos += Qty
                            l_Peso += Qty * itm.peso
                            l_Peso2 += Qty * itm.peso

                        Case "505"  'Mangas para PH
                            l_Mangas += Qty
                            l_Peso += Qty * itm.peso
                            l_Peso2 += Qty * itm.peso

                        Case "504"  'Mantenimiento hidrantes
                            l_EsTarea = True
                            l_Equipos += Qty

                        Case "551"  'Mantenimiento sistemas fijos
                            l_EsTarea = True
                            l_Equipos += Qty

                        Case "553"  'Relevamientos
                            l_EsTarea = True

                            If dr("itmref_0").ToString = "553010" Then
                                l_Mangas += Qty

                            Else
                                l_Equipos += Qty

                            End If

                        Case "652", "659"   'Controles y visitas / Otros
                            l_EsTarea = True
                            If l_Equipos = 0 Then l_Equipos += Qty

                    End Select

                Case 2  'Solapa préstamos
                    If itm.Categoria = "60" Then

                        If itm.Codigo.StartsWith("60100") Then
                            l_PrestamosExt += Qty
                            l_Peso += Qty * itm.peso

                        ElseIf itm.Codigo.StartsWith("60700") Then
                            l_PrestamosMan += Qty
                            l_Peso += Qty * itm.peso

                        End If

                    End If

            End Select

        Next

    End Sub
    Private Sub AnalizarIntervencionEntrega()
        Dim dr As DataRow
        Dim txt As String = ""
        Dim itm As New Articulo(cn)

        'Recorro todos los consumos cargados en el remito
        For Each dr In dt5.Rows
            Dim Qty As Integer = CInt(dr("hdtqty_0"))

            itm.Abrir(dr("hdtitm_0").ToString)

            l_Peso += Qty * itm.peso
            If itm.EsCarro Then l_TieneCarro = True

            If itm.LineaProducto = "453" Then 'Rechazo Extintores
                l_RechazosExt += Qty
                l_Peso2 += Qty * itm.peso

            ElseIf itm.LineaProducto = "503" Then 'Rechazo Mangueras
                l_RechazosMan += Qty
                l_Peso2 += Qty * itm.peso

            ElseIf itm.LineaProducto = "451" Then 'Extintores
                l_Equipos += Qty
                l_Peso2 += Qty * itm.peso

            ElseIf itm.LineaProducto = "505" Then 'Mangueras
                l_Mangas += Qty
                l_Peso2 += Qty * itm.peso

            ElseIf itm.Categoria = "60" Then

                If itm.Codigo.StartsWith("60100") Then
                    l_PrestamosExt += Qty

                ElseIf dr("hdtitm_0").ToString.StartsWith("60700") Then
                    l_PrestamosMan += Qty

                End If

            End If

        Next

    End Sub

    Public ReadOnly Property Equipos() As Integer Implements IRuteable.Equipos
        Get
            Return l_Equipos
        End Get
    End Property
    Public ReadOnly Property RechazosExtintor() As Integer Implements IRuteable.RechazosExtintor
        Get
            Return l_RechazosExt
        End Get
    End Property
    Public ReadOnly Property Mangueras() As Integer Implements IRuteable.Mangueras
        Get
            Return l_Mangas
        End Get
    End Property
    Public ReadOnly Property RechazosManguera() As Integer Implements IRuteable.RechazosManguera
        Get
            Return l_RechazosMan
        End Get
    End Property
    Public ReadOnly Property PrestamosExtintores() As Integer Implements IRuteable.PrestamosExtintores
        Get
            Return l_PrestamosExt
        End Get
    End Property
    Public ReadOnly Property PrestamosMangueras() As Integer Implements IRuteable.PrestamosMangueras
        Get
            Return l_PrestamosMan
        End Get
    End Property
    Public ReadOnly Property PesoUnigis() As Double Implements IRuteable.PesoUnigis
        Get
            Return l_Peso2
        End Get
    End Property
    Public ReadOnly Property TieneCarro() As Boolean Implements IRuteable.TieneCarro
        Get
            Return l_TieneCarro
        End Get
    End Property
    Public ReadOnly Property Pedido() As Pedido Implements IRuteable.Pedido
        Get
            Return Nothing
        End Get
    End Property
    Public ReadOnly Property TipoTarea() As String Implements IRuteable.TipoTarea
        Get
            If Remito.Trim <> "" OrElse Tipo = "G1" Then
                Return "ENT"

            Else
                Return IIf(l_EsTarea, "CTL", "RET").ToString

            End If

        End Get
    End Property
    Public Function ExisteArticulo(ByVal Articulo As String) As Boolean
        Dim flg As Boolean = False

        For Each dr As DataRow In dt2.Rows
            If dr("itmref_0").ToString = Articulo Then
                flg = True
                Exit For
            End If
        Next

        Return flg

    End Function
    Public ReadOnly Property Domicilio() As String Implements IRuteable.Domicilio
        Get
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                Dim dr As DataRow = dt1.Rows(0)
                txt = dr("add_0").ToString
            End If

            Return txt

        End Get
    End Property
    Public ReadOnly Property Localidad() As String Implements IRuteable.Localidad
        Get
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                Dim dr As DataRow = dt1.Rows(0)
                txt = dr("cty_0").ToString
            End If

            Return txt

        End Get
    End Property
    Public ReadOnly Property CodigoTercero() As String Implements IRuteable.CodigoTercero
        Get
            Return dt1.Rows(0).Item("bpc_0").ToString
        End Get
    End Property
    Public ReadOnly Property NombreTercero() As String Implements IRuteable.NombreTercero
        Get
            Return Me.Cliente.Nombre
        End Get
    End Property
    Public ReadOnly Property Instalaciones() As Integer Implements IRuteable.Instalaciones
        Get
            Return 0
        End Get
    End Property
    Public ReadOnly Property Cobranzas() As Boolean Implements IRuteable.Cobranza
        Get
            Dim flg As Boolean = False

            If Me.Remito.Trim <> "" Then
                Select Case Me.SolicitudAsociada.CondicionPagoCodigo 'dtSolicitud.Rows(0).Item("srepte_0").ToString
                    Case "001", "002"
                        flg = (Me.PrioridadEntrga = 1)
                End Select
            End If

            Return flg
        End Get
    End Property
    ReadOnly Property PrioridadEntrga() As Integer
        Get
            Dim x As Integer = 0

            Dim dr As DataRow = dt1.Rows(0)
            x = CInt(dr("dlvpio_0"))

            Return x
        End Get
    End Property
    Public ReadOnly Property Varios() As Boolean Implements IRuteable.Varios
        Get
            Return l_varios
        End Get
    End Property
    Public ReadOnly Property Hora() As String Implements IRuteable.Hora
        Get
            Dim txt As String = " "

            If Me.Franja1Desde <> "0000" And Me.Franja1Hasta <> "0000" Then
                txt = String.Format("{0} a {1}", Me.Franja1Desde, Me.Franja1Hasta)

                If Me.Franja2Desde <> "0000" And Me.Franja2Hasta <> "0000" Then
                    txt &= " y "
                    txt &= String.Format("{0} a {1}", Me.Franja2Desde, Me.Franja2Hasta)
                End If

            End If

            Return txt
        End Get
    End Property
    Public ReadOnly Property Peso() As Double Implements IRuteable.Peso
        Get

        End Get
    End Property

End Class