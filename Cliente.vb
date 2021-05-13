Imports System.Data.OracleClient
Imports System.Windows.Forms
Imports System.ComponentModel

Public Class Cliente
    Inherits Tercero
    Implements IDisposable

    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable
    Private disposedValue As Boolean = False        ' Para detectar llamadas redundantes

    '2015-01-15 TABLA XA8 - 
    'Desarrollo de Agenda de Precios Clientes
    Private da2 As OracleDataAdapter
    Private dt2 As New DataTable
    Private _Sucursales As SucursalCollection
    Private _Percepciones As Percepciones

    'NEW
    Public Sub New(ByVal cn As OracleConnection)
        MyBase.New(cn)
        Me.cn = cn
        Adaptadores()

        da.FillSchema(dt, SchemaType.Mapped)

    End Sub

    'SUB
    Private Sub Adaptadores()
        Dim Sql As String

        Sql = "SELECT * FROM bpcustomer WHERE bpcnum_0 = :bpcnum_0"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("bpcnum_0", OracleType.VarChar)
        da.InsertCommand = New OracleCommandBuilder(da).GetInsertCommand
        da.UpdateCommand = New OracleCommandBuilder(da).GetUpdateCommand

        '2015-01-15 TABLA XA8 - 
        'Desarrollo de Agenda de Precios Clientes
        Sql = "SELECT * FROM xa8 where bpcnum_0 = :bpcnum"
        da2 = New OracleDataAdapter(Sql, cn)
        da2.SelectCommand.Parameters.Add("bpcnum", OracleType.VarChar)

    End Sub
    Public Overloads Function Abrir(ByVal Codigo As String) As Boolean
        da.SelectCommand.Parameters("bpcnum_0").Value = Codigo
        dt.Clear()
        da.Fill(dt)

        MyBase.Abrir(Codigo)

        '2015-01-15 TABLA XA8 - 
        'Desarrollo de Agenda de Precios Clientes
        da2.SelectCommand.Parameters("bpcnum").Value = Codigo
        dt2.Clear()
        da2.Fill(dt2)

        _Sucursales = Nothing
        _Percepciones = Nothing

        Return dt.Rows.Count > 0

    End Function
    Public Overloads Sub Grabar()
        MyBase.Grabar()

        Dim Suc As String = ""

        'Grabo todas las sucursales del cliente
        For Each s As Sucursal In Me.Sucursales
            If s.SucursalPrincipal Then Suc = s.Sucursal
            s.Grabar()
        Next

        'Obtengo la sucursal factura si el cliente no la tiene establecida
        If Me.SucursalFactura.Trim = "" AndAlso Suc <> "" Then
            Me.SucursalFactura = Suc
        End If


        'Actualizo cliente
        da.Update(dt)

    End Sub
    Public Overloads Sub Nuevo(ByVal Tipo As Integer)
        MyBase.Nuevo()

        Dim dr As DataRow
        dr = dt.NewRow
        dr("BPCNUM_0") = " " 'numero
        dr("BPCNAM_0") = " " 'nombre
        dr("BPCSHO_0") = " "
        dr("BPCTYP_0") = 1
        dr("BPCINV_0") = " " 'numero
        dr("BPAINV_0") = " " '"001"
        dr("BPCPYR_0") = " " 'terceros
        dr("BPCGRU_0") = " " 'numero
        dr("BPCSTA_0") = 2
        dr("PPTFLG_0") = 1
        dr("BPCBPSNUM_0") = " "
        dr("DOCTYP_0") = " " 'doctip
        dr("FCTNUM_0") = " "
        dr("CHGTYP_0") = 3
        dr("DOCNUM_0") = " " 'dni
        dr("COMCAT_0") = 0 'comision
        dr("REP_0") = " " 'rep
        dr("REP_1") = " " 'rep
        dr("REP_2") = "11"
        dr("VACBPR_0") = " " 'iva
        dr("VATEXN_0") = " "
        dr("PTE_0") = " " 'pte
        dr("FREINV_0") = 1
        dr("DEP_0") = " "
        dr("INVDTAAMT_0") = 0
        dr("INVDTAAMT_1") = 0
        dr("INVDTAAMT_2") = 0
        dr("INVDTAAMT_3") = 0
        dr("INVDTAAMT_4") = 0
        dr("INVDTAAMT_5") = 0
        dr("INVDTAAMT_6") = 0
        dr("INVDTAAMT_7") = 0
        dr("INVDTAAMT_8") = 0
        dr("INVDTAAMT_9") = 0
        dr("INVDTA_0") = 10
        dr("INVDTA_1") = 0
        dr("INVDTA_2") = 0
        dr("INVDTA_3") = 0
        dr("INVDTA_4") = 0
        dr("INVDTA_5") = 0
        dr("INVDTA_6") = 0
        dr("INVDTA_7") = 0
        dr("INVDTA_8") = 0
        dr("INVDTA_9") = 0
        dr("TSCCOD_0") = " " 'fs1
        dr("TSCCOD_1") = " " 'fs2
        dr("TSCCOD_2") = " " 'fs3
        dr("TSCCOD_3") = " "
        dr("TSCCOD_4") = " "
        dr("PRITYP_0") = 1
        dr("BPCREM_0") = " " 'observa
        dr("OSTCTL_0") = 2
        dr("OSTAUZ_0") = 0
        dr("ORDMINAMT_0") = 0
        dr("CDTISR_0") = 0
        dr("CDTISRDAT_0") = #12/31/1599#
        dr("FUPTYP_0") = 1
        dr("FUPMINAMT_0") = 0
        dr("SOIPER_0") = 1
        dr("PAYBAN_0") = " "
        dr("ACCCOD_0") = "LOCAL"
        dr("CCE_0") = " "
        dr("CCE_1") = " "
        dr("CCE_2") = " "
        dr("CCE_3") = " "
        dr("CCE_4") = " "
        dr("CCE_5") = " "
        dr("CCE_6") = " "
        dr("CCE_7") = " "
        dr("CCE_8") = " "
        dr("MTCFLG_0") = 0
        dr("ORDTEX_0") = " "
        dr("INVTEX_0") = " "
        dr("LNDAUZ_0") = 2
        dr("OCNFLG_0") = 1
        dr("COPNBR_0") = 1
        dr("INVPER_0") = 1
        dr("DUDCLC_0") = 1
        dr("ORDCLE_0") = 2
        dr("ODL_0") = 1
        dr("DME_0") = 3
        dr("IME_0") = 6
        dr("BUS_0") = " "
        dr("ORIPPT_0") = " "
        dr("SALFCY_0") = " "
        dr("PITCDT_0") = 0
        dr("PITCPT_0") = 0
        dr("TOTPIT_0") = 0
        dr("COTCHX_0") = " "
        dr("COTPITRQD_0") = 0
        dr("CNTFIRDAT_0") = #12/31/1599#
        dr("ORDFIRDAT_0") = #12/31/1599#
        dr("QUOLASDAT_0") = #12/31/1599#
        dr("CNTLASDAT_0") = #12/31/1599#
        dr("CNTNEXDAT_0") = #12/31/1599#
        dr("CNTLASTYP_0") = 1
        dr("CNTNEXTYP_0") = 1
        dr("ABCCLS_0") = 6
        dr("AGTPCP_0") = 1
        dr("AGTSATTAX_0") = 1
        dr("SATTAX_0") = "BUE"
        dr("SATTAX_1") = "CBA"
        dr("SATTAX_2") = "CFE"
        dr("SATTAX_3") = "CHA"
        dr("SATTAX_4") = "CHU"
        dr("SATTAX_5") = "COR"
        dr("SATTAX_6") = "CTC"
        dr("SATTAX_7") = "ERI"
        dr("SATTAX_8") = "FMA"
        dr("SATTAX_9") = "JJY"
        dr("SATTAX_10") = "LPA"
        dr("SATTAX_11") = "LRJ"
        dr("SATTAX_12") = "MDZ"
        dr("SATTAX_13") = "MIS"
        dr("SATTAX_14") = "NQN"
        dr("SATTAX_15") = "RNG"
        dr("SATTAX_16") = "SCZ"
        dr("SATTAX_17") = "SDE"
        dr("SATTAX_18") = "SFE"
        dr("SATTAX_19") = "SJN"
        dr("SATTAX_20") = "SLA"
        dr("SATTAX_21") = "SLS"
        dr("SATTAX_22") = "TDF"
        dr("SATTAX_23") = "TUC"
        dr("SATTAX_24") = " "
        dr("FLGSATTAX_0") = 1
        dr("FLGSATTAX_1") = 1
        dr("FLGSATTAX_2") = 1
        dr("FLGSATTAX_3") = 1
        dr("FLGSATTAX_4") = 0
        dr("FLGSATTAX_5") = 0
        dr("FLGSATTAX_6") = 0
        dr("FLGSATTAX_7") = 0
        dr("FLGSATTAX_8") = 0
        dr("FLGSATTAX_9") = 0
        dr("FLGSATTAX_10") = 0
        dr("FLGSATTAX_11") = 0
        dr("FLGSATTAX_12") = 0
        dr("FLGSATTAX_13") = 0
        dr("FLGSATTAX_14") = 0
        dr("FLGSATTAX_15") = 0
        dr("FLGSATTAX_16") = 0
        dr("FLGSATTAX_17") = 0
        dr("FLGSATTAX_18") = 0
        dr("FLGSATTAX_19") = 0
        dr("FLGSATTAX_20") = 0
        dr("FLGSATTAX_21") = 0
        dr("FLGSATTAX_22") = 0
        dr("FLGSATTAX_23") = 0
        dr("FLGSATTAX_24") = 0
        dr("PORCSTATAX_0") = 0
        dr("PORCSTATAX_1") = 0
        dr("PORCSTATAX_2") = 0
        dr("PORCSTATAX_3") = 0
        dr("PORCSTATAX_4") = 0
        dr("PORCSTATAX_5") = 0
        dr("PORCSTATAX_6") = 0
        dr("PORCSTATAX_7") = 0
        dr("PORCSTATAX_8") = 0
        dr("PORCSTATAX_9") = 0
        dr("PORCSTATAX_10") = 0
        dr("PORCSTATAX_11") = 0
        dr("PORCSTATAX_12") = 0
        dr("PORCSTATAX_13") = 0
        dr("PORCSTATAX_14") = 0
        dr("PORCSTATAX_15") = 0
        dr("PORCSTATAX_16") = 0
        dr("PORCSTATAX_17") = 0
        dr("PORCSTATAX_18") = 0
        dr("PORCSTATAX_19") = 0
        dr("PORCSTATAX_20") = 0
        dr("PORCSTATAX_21") = 0
        dr("PORCSTATAX_22") = 0
        dr("PORCSTATAX_23") = 0
        dr("PORCSTATAX_24") = 0
        dr("EXPNUM_0") = 1
        dr("TIPDOC_0") = " "
        dr("NRODOC_0") = " "
        dr("CREUSR_0") = " "
        dr("CREDAT_0") = Date.Today 'Today.Hour * 3600 + Today.Minute * 60
        dr("UPDUSR_0") = " "
        dr("UPDDAT_0") = #12/31/1599# 'Date.Today 'Today.Hour * 3600 + Today.Minute * 60
        dr("XPROSTA_0") = 1 'Provincias Incluidas
        dr("XPROPER_0") = " "
        dr("XPROPER_1") = " "
        dr("XPROPER_2") = " "
        dr("XPROPER_3") = " "
        dr("XPROPER_4") = " "
        dr("XPROPER_5") = " "
        dr("XPROPER_6") = " "
        dr("XPROPER_7") = " "
        dr("XPROPER_8") = " "
        dr("XPROPER_9") = " "
        dr("XPROPER_10") = " "
        dr("XPROPER_11") = " "
        dr("XPROPER_12") = " "
        dr("XPROPER_13") = " "
        dr("XPROPER_14") = " "
        dr("XPROPER_15") = " "
        dr("XPROPER_16") = " "
        dr("XPROPER_17") = " "
        dr("XPROPER_18") = " "
        dr("XPROPER_19") = " "
        dr("XPROPER_20") = " "
        dr("XPROPER_21") = " "
        dr("XPROPER_22") = " "
        dr("XPROPER_23") = " "
        dr("XPROPER_24") = " "
        dr("XPROPER_25") = " "
        dr("XPROPER_26") = " "
        dr("XPROPER_27") = " "
        dr("XPROPER_28") = " "
        dr("XPROPER_29") = " "
        dr("XBPRPTH_0") = " "
        dr("XPRORAT_0") = 0
        dr("XPRORAT_1") = 0
        dr("XPRORAT_2") = 0
        dr("XPRORAT_3") = 0
        dr("XPRORAT_4") = 0
        dr("XPRORAT_5") = 0
        dr("XPRORAT_6") = 0
        dr("XPRORAT_7") = 0
        dr("XPRORAT_8") = 0
        dr("XPRORAT_9") = 0
        dr("XPRORAT_10") = 0
        dr("XPRORAT_11") = 0
        dr("XPRORAT_12") = 0
        dr("XPRORAT_13") = 0
        dr("XPRORAT_14") = 0
        dr("XPRORAT_15") = 0
        dr("XPRORAT_16") = 0
        dr("XPRORAT_17") = 0
        dr("XPRORAT_18") = 0
        dr("XPRORAT_19") = 0
        dr("XPRORAT_20") = 0
        dr("XPRORAT_21") = 0
        dr("XPRORAT_22") = 0
        dr("XPRORAT_23") = 0
        dr("XPRORAT_24") = 0
        dr("XPRORAT_25") = 0
        dr("XPRORAT_26") = 0
        dr("XPRORAT_27") = 0
        dr("XPRORAT_28") = 0
        dr("XPRORAT_29") = 0
        dr("XVAT_0") = " "
        dr("XVAT_1") = " "
        dr("XVAT_2") = " "
        dr("XVAT_3") = " "
        dr("XVAT_4") = " "
        dr("XVAT_5") = " "
        dr("XVAT_6") = " "
        dr("XVAT_7") = " "
        dr("XVAT_8") = " "
        dr("XVAT_9") = " "
        dr("XVAT_10") = " "
        dr("XVAT_11") = " "
        dr("XVAT_12") = " "
        dr("XVAT_13") = " "
        dr("XVAT_14") = " "
        dr("XVAT_15") = " "
        dr("XVAT_16") = " "
        dr("XVAT_17") = " "
        dr("XVAT_18") = " "
        dr("XVAT_19") = " "
        dr("XVAT_20") = " "
        dr("XVAT_21") = " "
        dr("XVAT_22") = " "
        dr("XVAT_23") = " "
        dr("XVAT_24") = " "
        dr("XVAT_25") = " "
        dr("XVAT_26") = " "
        dr("XVAT_27") = " "
        dr("XVAT_28") = " "
        dr("XVAT_29") = " "
        dr("XVATTYP_0") = 0
        dr("XVATTYP_1") = 0
        dr("XVATTYP_2") = 0
        dr("XVATTYP_3") = 0
        dr("XVATTYP_4") = 0
        dr("XVATTYP_5") = 0
        dr("XVATTYP_6") = 0
        dr("XVATTYP_7") = 0
        dr("XVATTYP_8") = 0
        dr("XVATTYP_9") = 0
        dr("XVATTYP_10") = 0
        dr("XVATTYP_11") = 0
        dr("XVATTYP_12") = 0
        dr("XVATTYP_13") = 0
        dr("XVATTYP_14") = 0
        dr("XVATTYP_15") = 0
        dr("XVATTYP_16") = 0
        dr("XVATTYP_17") = 0
        dr("XVATTYP_18") = 0
        dr("XVATTYP_19") = 0
        dr("XVATTYP_20") = 0
        dr("XVATTYP_21") = 0
        dr("XVATTYP_22") = 0
        dr("XVATTYP_23") = 0
        dr("XVATTYP_24") = 0
        dr("XVATTYP_25") = 0
        dr("XVATTYP_26") = 0
        dr("XVATTYP_27") = 0
        dr("XVATTYP_28") = 0
        dr("XVATTYP_29") = 0
        dr("XEXCERT_0") = " "
        dr("XEXCERT_1") = " "
        dr("XEXCERT_2") = " "
        dr("XEXCERT_3") = " "
        dr("XEXCERT_4") = " "
        dr("XEXCERT_5") = " "
        dr("XEXCERT_6") = " "
        dr("XEXCERT_7") = " "
        dr("XEXCERT_8") = " "
        dr("XEXCERT_9") = " "
        dr("XEXCERT_10") = " "
        dr("XEXCERT_11") = " "
        dr("XEXCERT_12") = " "
        dr("XEXCERT_13") = " "
        dr("XEXCERT_14") = " "
        dr("XEXCERT_15") = " "
        dr("XEXCERT_16") = " "
        dr("XEXCERT_17") = " "
        dr("XEXCERT_18") = " "
        dr("XEXCERT_19") = " "
        dr("XEXCERT_20") = " "
        dr("XEXCERT_21") = " "
        dr("XEXCERT_22") = " "
        dr("XEXCERT_23") = " "
        dr("XEXCERT_24") = " "
        dr("XEXCERT_25") = " "
        dr("XEXCERT_26") = " "
        dr("XEXCERT_27") = " "
        dr("XEXCERT_28") = " "
        dr("XEXCERT_29") = " "
        dr("XEXPORC_0") = 0
        dr("XEXPORC_1") = 0
        dr("XEXPORC_2") = 0
        dr("XEXPORC_3") = 0
        dr("XEXPORC_4") = 0
        dr("XEXPORC_5") = 0
        dr("XEXPORC_6") = 0
        dr("XEXPORC_7") = 0
        dr("XEXPORC_8") = 0
        dr("XEXPORC_9") = 0
        dr("XEXPORC_10") = 0
        dr("XEXPORC_11") = 0
        dr("XEXPORC_12") = 0
        dr("XEXPORC_13") = 0
        dr("XEXPORC_14") = 0
        dr("XEXPORC_15") = 0
        dr("XEXPORC_16") = 0
        dr("XEXPORC_17") = 0
        dr("XEXPORC_18") = 0
        dr("XEXPORC_19") = 0
        dr("XEXPORC_20") = 0
        dr("XEXPORC_21") = 0
        dr("XEXPORC_22") = 0
        dr("XEXPORC_23") = 0
        dr("XEXPORC_24") = 0
        dr("XEXPORC_25") = 0
        dr("XEXPORC_26") = 0
        dr("XEXPORC_27") = 0
        dr("XEXPORC_28") = 0
        dr("XEXPORC_29") = 0
        dr("XEXDAT_0") = #12/31/1599#
        dr("XEXDAT_1") = #12/31/1599#
        dr("XEXDAT_2") = #12/31/1599#
        dr("XEXDAT_3") = #12/31/1599#
        dr("XEXDAT_4") = #12/31/1599#
        dr("XEXDAT_5") = #12/31/1599#
        dr("XEXDAT_6") = #12/31/1599#
        dr("XEXDAT_7") = #12/31/1599#
        dr("XEXDAT_8") = #12/31/1599#
        dr("XEXDAT_9") = #12/31/1599#
        dr("XEXDAT_10") = #12/31/1599#
        dr("XEXDAT_11") = #12/31/1599#
        dr("XEXDAT_12") = #12/31/1599#
        dr("XEXDAT_13") = #12/31/1599#
        dr("XEXDAT_14") = #12/31/1599#
        dr("XEXDAT_15") = #12/31/1599#
        dr("XEXDAT_16") = #12/31/1599#
        dr("XEXDAT_17") = #12/31/1599#
        dr("XEXDAT_18") = #12/31/1599#
        dr("XEXDAT_19") = #12/31/1599#
        dr("XEXDAT_20") = #12/31/1599#
        dr("XEXDAT_21") = #12/31/1599#
        dr("XEXDAT_22") = #12/31/1599#
        dr("XEXDAT_23") = #12/31/1599#
        dr("XEXDAT_24") = #12/31/1599#
        dr("XEXDAT_25") = #12/31/1599#
        dr("XEXDAT_26") = #12/31/1599#
        dr("XEXDAT_27") = #12/31/1599#
        dr("XEXDAT_28") = #12/31/1599#
        dr("XEXDAT_29") = #12/31/1599#
        dr("XCSHCAT_0") = " "
        dr("XCSHDAYPAY_0") = 0
        dr("XBLQSOH_0") = 0
        dr("XBLQMOTIV_0") = 0
        dr("YBPRRAT_0") = 0
        dr("XBLQSOHMIN_0") = 0
        dr("XRATBPR_0") = 0
        dr("XBLQTOL_0") = 0
        dr("XRATBPRTYP_0") = 1
        dr("YBPRATYP_0") = 0
        dr("ZONA_0") = " "
        dr("XRI_0") = 1
        dr("XRIF_0") = #12/31/1599#
        dr("XDI_0") = 1
        dr("XHI_0") = 1
        dr("XIF_0") = 1
        dr("XREQAUT_0") = 1
        dr("XABO_0") = 1
        dr("XPQ_0") = 1
        dr("XCO2_0") = 1
        dr("XESPUMA_0") = 1
        dr("XHALOG_0") = 1
        dr("XACETA_0") = 1
        dr("XPQE_0") = 1
        dr("XFCRTO_0") = 1
        dr("XLSTESP_0") = 1
        dr("XMAILFC_0") = " " 'MailFC
        dr("XMAILFCFLG_0") = 2 'IIf(pte < "022", 1, 2)
        dr("X415_0") = 1
        dr("XOC_0") = 1
        dr("XWEBPWD_0") = " "
        dr("XWEBDAT_0") = #12/31/1599#
        dr("BPCLON_0") = " " 'fantasia
        dr("XRECI_0") = 1
        dr("XREQFAC_0") = 1
        dr("XACTIV_0") = 0
        dr("CTRLFLG_0") = 2
        'Agregado por adonix
        dr("xregib_0") = 0
        dr("xregib_1") = 0
        dr("xregib_2") = 0
        dr("xregib_3") = 0
        dr("xregib_4") = 0
        dr("xregib_5") = 0
        dr("xregib_6") = 0
        dr("xregib_7") = 0
        dr("xregib_8") = 0
        dr("xregib_9") = 0
        dr("xregib_10") = 0
        dr("xregib_11") = 0
        dr("xregib_12") = 0
        dr("xregib_13") = 0
        dr("xregib_14") = 0
        dr("xregib_15") = 0
        dr("xregib_16") = 0
        dr("xregib_17") = 0
        dr("xregib_18") = 0
        dr("xregib_19") = 0
        dr("xregib_20") = 0
        dr("xregib_21") = 0
        dr("xregib_22") = 0
        dr("xregib_23") = 0
        dr("xregib_24") = 0
        dr("xregib_25") = 0
        dr("xregib_26") = 0
        dr("xregib_27") = 0
        dr("xregib_28") = 0
        dr("xregib_29") = 0
        dr("xcondiibb_0") = " "
        dr("xflgar_0") = 1
        dr("xibcfergsi_0") = 1
        dr("xvtaantes_0") = 1
        dr("xcpyfac_0") = " "
        dr("XPOTENCIAL_0") = 1
        dr("MAILMKT_0") = 2
        dr("MAILCOB_0") = 2
        dr("MAILVTA_0") = 2
        dr("XPINTURA_0") = 1
        dr("ABCPLUS_0") = 1
        dr("XCOMPLEJ_0") = " "

        dt.Rows.Add(dr)

        If Tipo = 1 Then ConvetirEnCliente()

    End Sub
    Public Sub ConvetirEnCliente()
        Dim dr As DataRow

        Tipo = 1
        EsCliente = True
        EsProspecto = False

        dr = dt.Rows(0)
        dr.BeginEdit()
        dr("ordcle_0") = 2
        dr("lndauz_0") = 2
        dr("dme_0") = 3
        dr("freinv_0") = 1
        dr("ime_0") = 6
        dr("invper_0") = 1
        dr("ostctl_0") = 2
        dr("acccod_0") = "LOCAL"
        dr("fuptyp_0") = 1
        dr("soiper_0") = 1

        dr.EndEdit()
    End Sub
    'FUNCTION
    Public Function CheckCuit() As Boolean
        Return CheckCuit(Me.CUIT.Trim)
    End Function
    Shared Function CheckCuit(ByVal Nro As String) As Boolean
        Dim cuit As String = Nro
        Dim cuit2 As String = ""
        Dim n(9) As Integer
        Dim v1 As Integer = 0
        Dim v2 As Integer = 0
        Dim v3 As Integer = 0

        n(0) = 5 : n(1) = 4 : n(2) = 3 : n(3) = 2 : n(4) = 7
        n(5) = 6 : n(6) = 5 : n(7) = 4 : n(8) = 3 : n(9) = 2

        If cuit.Length <> 11 Then cuit = ""

        For i As Integer = 0 To cuit.Length - 2
            Dim c As String

            c = cuit.Substring(i, 1)

            If IsNumeric(c) Then
                Dim j As Integer = CInt(c)

                cuit2 &= c
                v1 += j * n(i)

            Else
                Exit For

            End If

        Next

        v2 = v1 Mod 11
        v3 = 11 - v2

        Select Case v3
            Case 11
                cuit2 &= "0"
            Case 10
                cuit2 &= "9"
            Case Else
                cuit2 &= v3
        End Select

        Return cuit = cuit2

    End Function
    Shared Function CheckDni(ByVal Nro As String) As Boolean
        Dim l As Integer = 0
        Dim flg As Boolean = True

        For i As Integer = 0 To Nro.Length - 1
            If IsNumeric(Nro.Substring(i, 1)) Then
                l += 1
            Else
                flg = False
            End If
        Next

        Return ((l = 7 Or l = 8) And flg)

    End Function
    Public Function EmpresaFacturacion() As String
        'Consulta por que empresa se hizo la última Solicitud de servicio
        Dim Sql As String
        Dim dt As New DataTable
        Dim da As OracleDataAdapter
        Dim dr As DataRow

        Sql = "SELECT srenum_0, num_0, sre.credat_0 "
        Sql &= "FROM serrequest sre INNER JOIN interven itn ON (srenum_0 = srvdemnum_0) "
        Sql &= "WHERE srebpc_0 = :srebpc_0 AND typ_0 <> 'F1' "
        Sql &= "ORDER BY sre.credat_0 DESC"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("srebpc_0", OracleType.VarChar).Value = Codigo

        da.Fill(dt)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            If dr("srenum_0").ToString.StartsWith("D") Then
                EmpresaFacturacion = "DNY"
            Else
                EmpresaFacturacion = "MON"
            End If
        Else
            EmpresaFacturacion = " "

        End If

        dt.Dispose()
        da.Dispose()

    End Function
    Public Sub EnlazarCombo(ByVal cbo As ComboBox)
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim Sql As String

        Sql = "SELECT bpaadd_0, bpaadd_0 || ' - ' || bpaaddlig_0 || ' (' || cty_0 ||')' AS descripcion "
        Sql &= "FROM  bpaddress "
        Sql &= "WHERE bpanum_0 = :bpanum_0 "
        Sql &= "ORDER BY bpaadd_0"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("bpanum_0", OracleType.VarChar).Value = Codigo
        da.Fill(dt)

        With cbo
            .DataSource = dt
            .DisplayMember = "descripcion"
            .ValueMember = "bpaadd_0"
        End With

    End Sub
    Public Function SucursalesTabla() As DataTable
        Dim Sql = "SELECT bpaadd_0, bpaaddlig_0, poscod_0, cty_0, cry_0, sat_0, tel_0, bpaadd_0 || '-' || bpaaddlig_0 || ' (' || cty_0 || ')' AS direccion FROM bpaddress WHERE bpanum_0 = :bpanum_0 ORDER BY bpaadd_0"
        Dim da As New OracleDataAdapter(Sql, cn)
        Dim dt As New DataTable

        da.SelectCommand.Parameters.Add("bpanum_0", OracleType.VarChar).Value = Codigo
        da.Fill(dt)

        da.Dispose()

        Return dt

    End Function
    Public ReadOnly Property Sucursales() As SucursalCollection
        Get
            If _Sucursales Is Nothing Then
                _Sucursales = New SucursalCollection(cn)
                _Sucursales.CargarSucursales(Me.Codigo)
            End If

            Return _Sucursales
        End Get
    End Property
    Public Function TieneDatos() As Boolean
        Return dt.Rows.Count > 0
    End Function
    Public Function SolicitudesAbiertas() As DataTable
        'CREACION DE ADAPTADORES PARA: DASSOPEN
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable

        Sql = "SELECT srenum_0, srenum_0 || ' - ' || sreresdat_0 || ' - ' || conspt_0 AS texto FROM serrequest WHERE srebpc_0 = :srebpc_0 AND sreass_0 = 2 ORDER BY srenum_0 DESC"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand = New OracleCommand(Sql, cn)
        da.SelectCommand.Parameters.Add("srebpc_0", OracleType.VarChar).Value = Codigo
        da.Fill(dt)
        da.Dispose()
        Return dt

    End Function
    Public Function ContratoValidez() As Date
        'Función que indica si el cliente tiene un contrato de servicio vigente
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As DataTable

        'daAbonados
        Sql = "SELECT conenddat_0 "
        Sql &= "FROM contserv "
        Sql &= "WHERE fddflg_0 <> 2 AND rsiflg_0 <> 2 AND conbpc_0 = :conbpc_0 AND conenddat_0 >= to_date(:dat_0, 'dd/mm/yyyy') "
        Sql &= "ORDER BY conenddat_0"

        da = New OracleDataAdapter(Sql, cn)

        With da.SelectCommand.Parameters
            .Add("conbpc_0", OracleType.VarChar).Value = Codigo
            .Add("dat_0", OracleType.VarChar).Value = Today.ToString("dd/MM/yyyy")
        End With

        dt = New DataTable
        da.Fill(dt)
        da.Dispose()

        If dt.Rows.Count = 0 Then
            Return Nothing

        Else
            Return CType(dt.Rows(0).Item(0), Date)

        End If

    End Function
    Public Function AccesoUsuario(ByVal Usr As String) As Boolean
        Dim flg As Boolean = False

        '1.- Verificación si el cliente pertenece al vendedor
        If Vendedor.Codigo = Usr Then flg = True

        '2.- Verificacion si el analista tiene acceso al cliente
        If Vendedor.Analista.Codigo = Usr Then flg = True

        '3.- Verificacion si el usuario tiene permiso segun la configuracion de vencimientos
        If Not flg Then
            Dim da As OracleDataAdapter
            Dim dt As New DataTable
            Dim dr As DataRow
            Dim Sql As String

            Sql = "SELECT * FROM xnetvenc WHERE usr_0 = :usr_0 AND rep_0 = :rep_0"
            da = New OracleDataAdapter(Sql, cn)
            da.SelectCommand.Parameters.Add("usr_0", OracleType.VarChar).Value = Usr
            da.SelectCommand.Parameters.Add("rep_0", OracleType.VarChar).Value = Representante(0) '.Codigo
            da.Fill(dt)
            da.Dispose()

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)

                If EsAbonado And CInt(dr("abo_0")) = 1 Then flg = True
                If Not EsAbonado And CInt(dr("noabo_0")) = 1 Then flg = True
            End If

        End If

        Return flg

    End Function
    Public Function ExisteCuit(ByVal cuit As String) As DataTable
        Dim dt As New DataTable
        Dim da As OracleDataAdapter
        Dim sql As String

        sql = "select * "
        sql &= "from bpartner pb inner join "
        sql &= "     bpcustomer bpc on (pb.bprnum_0 = bpc.bpcnum_0) "
        sql &= "where crn_0 = :doc"

        da = New OracleDataAdapter(sql, cn)
        da.SelectCommand.Parameters.Add("doc", OracleType.VarChar).Value = cuit
        da.Fill(dt)

        Return dt

    End Function
    Public Function UltimaFactura() As Factura
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim Sql As String
        Dim sih As Factura = Nothing

        Sql = "select num_0 "
        Sql &= "from sinvoice "
        Sql &= "where bpr_0 = :bpr and "
        Sql &= "	  sivtyp_0 = 'FAC'"
        Sql &= "order by accdat_0 desc"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("bpr", OracleType.VarChar).Value = Me.Codigo
        da.Fill(dt)
        da.Dispose()

        If dt.Rows.Count > 0 Then
            Dim dr As DataRow = dt.Rows(0)
            sih = New Factura(cn, dr(0).ToString)
        End If

        Return sih

    End Function
    Public Function TieneFacturas() As Boolean
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim Sql As String
        Dim flg As Boolean = False

        Sql = "select num_0 "
        Sql &= "from sinvoice "
        Sql &= "where bpr_0 = :bpr and "
        Sql &= "	  sivtyp_0 = 'FAC'"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("bpr", OracleType.VarChar).Value = Me.Codigo
        da.Fill(dt)
        da.Dispose()

        flg = dt.Rows.Count > 0
        dt.Dispose()

        Return flg
    End Function
    Friend Sub AgregarPercepcion(ByVal Indice As Integer, ByVal CodigoPercepcion As String, ByVal Alicuota As Double)
        Dim dr As DataRow

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            dr.BeginEdit()
            dr("XPROPER_" & Indice.ToString) = CodigoPercepcion
            dr("XPRORAT_" & Indice.ToString) = Alicuota
            dr("XREGIB_" & Indice.ToString) = 4
            dr.EndEdit()
        End If

    End Sub
    Shared Function Existe(ByVal cn As OracleConnection, ByVal Codigo As String) As Boolean
        Dim da As OracleDataAdapter
        Dim Sql As String = ""
        Dim dt As New DataTable

        Sql = "SELECT bpcnum_0 "
        Sql &= "FROM bpcustomer "
        Sql &= "WHERE bpcnum_0 = :bpcnum"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("bpcnum", OracleType.VarChar).Value = Codigo
        da.Fill(dt)
        da.Dispose()

        Return dt.Rows.Count > 0

    End Function

    'PROPERTY
    Public Overloads Property Codigo() As String
        Get
            Dim dr As DataRow
            Dim bpcnum As String = " "

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                bpcnum = dr("bpcnum_0").ToString
            End If

            Return bpcnum

        End Get
        Set(ByVal value As String)
            If dt.Rows.Count > 0 Then
                Dim dr As DataRow
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("bpcnum_0") = IIf(value.Trim = "", " ", value.ToUpper)
                dr.EndEdit()
                MyBase.Codigo = value
            End If
        End Set
    End Property
    Public Property TerceroFactura() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("bpcinv_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dt.Rows(0)
            dr.BeginEdit()
            dr("bpcinv_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr.EndEdit()
        End Set
    End Property
    Public Property SucursalFactura() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("bpainv_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dt.Rows(0)
            dr.BeginEdit()
            dr("bpainv_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr.EndEdit()
        End Set
    End Property
    Public Property TerceroGrupoCodigo() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("bpcgru_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("bpcgru_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property TerceroGrupo() As Cliente
        Get
            Dim dr As DataRow = dt.Rows(0)
            Dim bpc As New Cliente(cn)
            bpc.Abrir(dr("bpcpyr_0").ToString)
            Return bpc
        End Get

    End Property
    Public ReadOnly Property TerceroPagador() As Cliente
        Get
            Dim dr As DataRow = dt.Rows(0)
            Dim bpc As New Cliente(cn)
            bpc.Abrir(TerceroPagadorCodigo)
            Return bpc
        End Get
    End Property
    Public ReadOnly Property TerceroPagadorCodigo() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("bpcpyr_0").ToString
        End Get
    End Property
    Public Overloads Property Nombre() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("bpcnam_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("bpcnam_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr.EndEdit()

            MyBase.Nombre = value
        End Set
    End Property
    Public Property Pagador() As String
        Get
            Dim dr As DataRow
            Dim txt As String = " "

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                txt = dr("bpcpyr_0").ToString
            End If
            Return txt
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("bpcpyr_0") = IIf(value.Trim = "", " ", value.ToUpper)
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property Tipo() As Integer
        Get
            Dim dr As DataRow
            Dim i As Integer = 0

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                i = CInt(dr("bpctyp_0"))
            End If
            Return i
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("bpctyp_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property Representante(ByVal Indice As Integer) As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("rep_" & Indice.ToString).ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)

            dr.BeginEdit()
            dr("rep_" & Indice.ToString) = value
            dr.EndEdit()

        End Set
    End Property
    Public Property RegimenImpuesto() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Dim txt As String = "CF"

            If dr("vacbpr_0").ToString() <> " " Then txt = dr("vacbpr_0").ToString()

            Return txt

        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("vacbpr_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property TipoCambio() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CInt(dr("chgtyp_0"))
        End Get
    End Property
    Public ReadOnly Property TipoPrecio() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CInt(dr("prityp_0"))
        End Get
    End Property
    Public ReadOnly Property CondicionPago() As CondicionPago
        Get
            Dim dr As DataRow
            Dim pte As String = ""

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                pte = dr("pte_0").ToString
            End If

            Return New CondicionPago(cn, pte)
        End Get
    End Property
    Public Property CondicionDePago() As String
        Get
            Dim dr As DataRow
            Dim pte As String = ""

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                pte = dr("pte_0").ToString
            End If
            Return pte
        End Get
        Set(ByVal value As String)
            If dt.Rows.Count > 0 Then
                Dim dr As DataRow
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("pte_0") = IIf(value.Trim = "", " ", value.ToUpper)
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property EsAbonado() As Boolean
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return (CInt(dr("xabo_0")) = 2)
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("xabo_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property NombreFantasia() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("bpclon_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            If value.Trim = "" Then value = " "
            dr = dt.Rows(0)
            dr.BeginEdit()
            dr("bpclon_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr.EndEdit()
        End Set
    End Property
    Public Property ActivoControlPeriodico() As Boolean
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CBool(IIf(CInt(dr("xactiv_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow = dt.Rows(0)
            dr = dt.Rows(0)
            dr.BeginEdit()
            dr("xactiv_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property ContratoActivo(ByVal Fecha As Date) As Boolean
        Get
            'Función que indica si el cliente tiene un contrato de servicio vigente
            Dim Sql As String
            Dim da1 As OracleDataAdapter
            Dim dt1 As DataTable

            'daAbonados
            Sql = "SELECT connum_0 " & _
                  "FROM contserv " & _
                  "WHERE fddflg_0 <> 2 AND " & _
                  "      rsiflg_0 <> 2 AND " & _
                  "      conbpc_0 = :p1 AND " & _
                  "      :p2 between constrdat_0 and conenddat_0"

            da1 = New OracleDataAdapter(Sql, cn)

            With da1.SelectCommand.Parameters
                .Add("p1", OracleType.VarChar).Value = Codigo
                .Add("p2", OracleType.DateTime).Value = Fecha
            End With

            dt1 = New DataTable
            da1.Fill(dt1)

            ContratoActivo = (dt1.Rows.Count > 0)

            da1.Dispose()
            dt1.Dispose()

        End Get
    End Property
    Public ReadOnly Property facturaAño(ByVal cliente As String) As Boolean
        Get
            Dim Sql As String
            Dim da1 As OracleDataAdapter
            Dim dt1 As DataTable
            Sql = "select bpr_0 from sinvoice  "
            Sql &= "where accdat_0 > :dat and bpr_0 = :bpc"
            da1 = New OracleDataAdapter(Sql, cn)
            Dim dia As Date = Today.AddMonths(-13)
            da1.SelectCommand.Parameters.Add("dat", OracleType.DateTime).Value = dia
            da1.SelectCommand.Parameters.Add("bpc", OracleType.VarChar).Value = cliente
            dt1 = New DataTable
            da1.Fill(dt1)

            facturaAño = (dt1.Rows.Count > 0)

            da1.Dispose()
            dt1.Dispose()
        End Get
    End Property
    Public ReadOnly Property intervencionAño(ByVal cliente As String) As Boolean
        Get
            Dim Sql As String
            Dim da1 As OracleDataAdapter
            Dim dt1 As DataTable
            Sql = "select bpc_0 from interven "
            Sql &= "where dat_0 > :dat and bpc_0 = :bpc"
            da1 = New OracleDataAdapter(Sql, cn)
            Dim dia As Date = Today.AddMonths(-12)
            da1.SelectCommand.Parameters.Add("dat", OracleType.DateTime).Value = dia
            da1.SelectCommand.Parameters.Add("bpc", OracleType.VarChar).Value = cliente

            dt1 = New DataTable
            da1.Fill(dt1)

            intervencionAño = (dt1.Rows.Count > 0)

            da1.Dispose()
            dt1.Dispose()
        End Get
    End Property
    Public Property Activo() As Boolean
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return (CInt(dr("bpcsta_0")) = 2)
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("bpcsta_0") = IIf(value, 2, 1)
                dr.EndEdit()
            End If
        End Set
    End Property
    Public ReadOnly Property Bloqueado() As Boolean
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return (CInt(dr("ostctl_0")) = 3)
        End Get
    End Property
    Public Property RequiereFacturaFisica() As Boolean
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CInt(dr("xreqfac_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("xreqfac_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Observaciones() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("bpcrem_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("bpcrem_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr.EndEdit()
        End Set
    End Property
    Public Property Vendedor1Codigo() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("rep_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("rep_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr.EndEdit()
        End Set
    End Property
    Public Property Vendedor2Codigo() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("rep_1").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("rep_1") = IIf(value.Trim = "", " ", value.ToUpper)
            dr.EndEdit()
        End Set
    End Property
    Public Property Vendedor3Codigo() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("rep_2").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("rep_2") = IIf(value.Trim = "", " ", value.ToUpper)
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property Vendedor() As Vendedor
        Get
            Dim dr As DataRow
            Dim rep As String = " "

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                rep = dr("rep_0").ToString
            End If

            Return New Vendedor(cn, rep)
        End Get
    End Property
    Public ReadOnly Property Vendedor2() As Vendedor
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return New Vendedor(cn, dr("rep_1").ToString)
        End Get
    End Property
    Public ReadOnly Property Vendedor3() As Vendedor
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return New Vendedor(cn, dr("rep_2").ToString)
        End Get
    End Property
    Public ReadOnly Property SucursalLegal() As Sucursal
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Dim suc As New Sucursal(cn, Codigo, dr("bpainv_0").ToString)

                Return suc

            Else
                Return Nothing

            End If

        End Get
    End Property
    Public ReadOnly Property SucursalEnvioFactura() As Sucursal
        Get
            Dim da As OracleDataAdapter
            Dim dt As New DataTable
            Dim Sql As String = ""
            Dim dr As DataRow
            Dim suc As Sucursal = Nothing

            Sql = "SELECT * "
            Sql &= "FROM bpaddress "
            Sql &= "WHERE bpanum_0 = :bpanum_0 AND xentfac_0 = 2"

            da = New OracleDataAdapter(Sql, cn)
            da.SelectCommand.Parameters.Add("bpanum_0", OracleType.VarChar).Value = Codigo

            da.Fill(dt)

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)

                suc = New Sucursal(cn, Codigo, dr("bpaadd_0").ToString)

            End If

            Return suc

        End Get
    End Property
    Public ReadOnly Property Honorario() As Double
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CDbl(dr("xratbpr_0"))
        End Get
    End Property
    Public Property Familia1() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("tsccod_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("tsccod_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr.EndEdit()
        End Set
    End Property
    Public Property Familia2() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("tsccod_1").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("tsccod_1") = IIf(value.Trim = "", " ", value.ToUpper)
            dr.EndEdit()
        End Set
    End Property
    Public Property Familia3() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("tsccod_2").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("tsccod_2") = IIf(value.Trim = "", " ", value.ToUpper)
            dr.EndEdit()
        End Set
    End Property
    Public Property Familia4() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("tsccod_3").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("tsccod_3") = IIf(value.Trim = "", " ", value.ToUpper)
            dr.EndEdit()
        End Set
    End Property
    Public Property Familia5() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("tsccod_4").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("tsccod_4") = IIf(value.Trim = "", " ", value.ToUpper)
            dr.EndEdit()
        End Set
    End Property
    Public Overloads Property TipoDoc() As String
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return dr("doctyp_0").ToString
            Else
                Return " "
            End If
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            MyBase.TipoDoc = value

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("doctyp_0") = IIf(value.Trim = "", " ", value.ToUpper)
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Overloads Property CUIT() As String
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return dr("docnum_0").ToString.Trim
            Else
                Return ""
            End If
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("docnum_0") = IIf(value.Trim = "", " ", value.ToUpper)
                dr.EndEdit()

                MyBase.CUIT = value
            End If
        End Set
    End Property
    Public Property MailFC() As String
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return dr("xmailfc_0").ToString.Trim
            Else
                Return ""
            End If
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("xmailfc_0") = IIf(value.Trim = "", " ", value)
                dr.EndEdit()
            End If
        End Set
    End Property
    Public ReadOnly Property EnviarMailFC() As Boolean
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return CInt(dr("xmailfcflg_0")) = 2
            Else
                Return False
            End If

        End Get
    End Property
    Public Property B2BAcceso() As Date
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return CDate(dr("xwebdat_0"))

            Else
                Return #12/31/1599#

            End If

        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("xwebdat_0") = value.Date
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property ClaveB2B() As String
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return dr("xwebpwd_0").ToString.Trim
            Else
                Return ""
            End If

        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If value.Trim = "" Then value = " "

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("xwebpwd_0") = IIf(value.Trim = "", " ", value.ToUpper)
                dr.EndEdit()
            End If

        End Set
    End Property
    Public ReadOnly Property Saldo(Optional ByVal Sociedad As String = "") As Double
        Get
            Dim da As OracleDataAdapter
            Dim dt As New DataTable
            Dim Sql As String
            Dim i As Double = 0
            Dim dr As DataRow

            If Sociedad = "" Then
                Sql = "SELECT SUM(sns_0 * (amtloc_0 - payloc_0)) AS saldo "
                Sql &= "FROM gaccdudate "
                Sql &= "WHERE flgcle_0 <> 2 AND typ_0 NOT IN ('APG', 'CLO') AND sac_0 = 'DVL' AND bpr_0 = :bpr_0"

                da = New OracleDataAdapter(Sql, cn)

            Else
                Sql = "SELECT SUM(sns_0 * (amtloc_0 - payloc_0)) AS saldo "
                Sql &= "FROM gaccdudate "
                Sql &= "WHERE flgcle_0 <> 2 AND typ_0 NOT IN ('APG', 'CLO') AND sac_0 = 'DVL' AND bpr_0 = :bpr_0 AND cpy_0 = :cpy_0"

                da = New OracleDataAdapter(Sql, cn)
                da.SelectCommand.Parameters.Add("cpy_0", OracleType.VarChar).Value = Sociedad

            End If

            da.SelectCommand.Parameters.Add("bpr_0", OracleType.VarChar).Value = Codigo

            Try
                da.Fill(dt)
                da.Dispose()

                dr = dt.Rows(0)

                i = CDbl(dr(0))

            Catch ex As Exception

            End Try

            Return i

        End Get
    End Property
    Public ReadOnly Property FcRto() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)

            Return CInt(IIf(CInt(dr("xfcrto_0")) = 2, 2, 1))

        End Get
    End Property
    Public Property CategoriaComision() As Integer
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return CInt(dr("comcat_0"))
            Else
                Return 0
            End If
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("comcat_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public ReadOnly Property TieneParque(ByVal sucursal As String) As Boolean
        Get
            Dim Sql As String
            Dim da1 As OracleDataAdapter
            Dim dt1 As DataTable
            Sql = "select * from machines where bpcnum_0 = :cliente and fcyitn_0 = :sucursal "
            da1 = New OracleDataAdapter(Sql, cn)
            da1.SelectCommand.Parameters.Add("cliente", OracleType.VarChar).Value = Codigo
            da1.SelectCommand.Parameters.Add("sucursal", OracleType.VarChar).Value = sucursal
            dt1 = New DataTable
            da1.Fill(dt1)

            TieneParque = (dt1.Rows.Count > 0)

            da1.Dispose()
            dt1.Dispose()
        End Get
    End Property
    Public Property QuiereControles() As Boolean
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return CInt(dr("ctrlflg_0")) = 2

            Else
                Return False

            End If

        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("xwebpwd_0") = IIf(value, 2, 1)
                dr.EndEdit()
            End If

        End Set
    End Property
    Public ReadOnly Property OC_obligatoria() As Boolean
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CBool(IIf(CInt(dr("xoc_0")) = 2, True, False))
        End Get
    End Property
    Public Property CondicionIb() As String
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return dr("xcondiibb_0").ToString.Trim
            Else
                Return ""
            End If
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("xcondiibb_0") = IIf(value.Trim = "", " ", value.ToUpper)
                dr.EndEdit()
            End If
        End Set
    End Property
    Public ReadOnly Property Percepciones() As Percepciones
        Get
            If _Percepciones Is Nothing Then
                _Percepciones = New Percepciones(cn, Me)
            End If

            Return _Percepciones
        End Get

    End Property
    Public Overloads Property EsProspecto() As Boolean
        Get
            Dim dr As DataRow
            Dim flg As Boolean = False

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)

                flg = CInt(dr("bpctyp_0")) = 4
            End If

            Return flg

        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("bpctyp_0") = IIf(value, 4, 1)
                dr.EndEdit()

                MyBase.EsProspecto = value

            End If

        End Set
    End Property
    Public ReadOnly Property MailMkt() As Boolean
        Get
            Dim dr As DataRow
            dr = dt.Rows(0)
            Return CBool(IIf(CInt(dr("MAILMKT_0")) = 2, True, False))

        End Get

    End Property
    Public ReadOnly Property MailCob() As Boolean
        Get
            Dim dr As DataRow
            dr = dt.Rows(0)
            Return CBool(IIf(CInt(dr("MAILCOB_0")) = 2, True, False))
        End Get

    End Property
    Public ReadOnly Property MailVta() As Boolean
        Get
            Dim dr As DataRow
            dr = dt.Rows(0)
            Return CBool(IIf(CInt(dr("MAILVTA_0")) = 2, True, False))
        End Get

    End Property
    Public Property AVentasAntesEntregar() As Boolean
        Get
            Dim dr As DataRow
            Dim flg As Boolean = False

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                flg = CInt(dr("xvtaantes_0")) = 2
            End If

            Return flg
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("xvtasantes_0") = IIf(value, 2, 1)
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property EmpresaFacturacionObligatoria() As String
        Get
            Dim dr As DataRow
            Dim v As String = ""

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                v = dr("xcpyfac_0").ToString.Trim
            End If

            Return v
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("xcpyfac_0") = IIf(value.Trim = "", " ", value)
                dr.EndEdit()
            End If
        End Set
    End Property

    Public Property XPQ() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CInt(dr("xpq_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("xpq_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property XPQE() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CInt(dr("xpqe_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("xpqe_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property xco2() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CInt(dr("xco2_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("xco2_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property XESPUMA() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CInt(dr("xespuma_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("xespuma_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property XHALOG() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CInt(dr("xhalog_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("xhalog_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property XACETA() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CInt(dr("xaceta_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("xaceta_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property XHI() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CInt(dr("xhi_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("xhi_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public ReadOnly Property TipoAbcStr() As String
        Get
            Dim i As Integer
            Dim s As String = ""

            i = TipoAbc

            If i > 0 And i <= 6 Then
                s = Chr(64 + i)
                If i = 1 AndAlso Me.PlusAbc Then
                    s &= "+"
                End If
            End If

            Return s

        End Get
    End Property

    Public Property TipoAbc() As Integer
        Get
            Dim dr As DataRow
            Dim i As Integer = 0

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                i = CInt(dr("abccls_0"))
            End If

            Return i
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("abccls_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property PlusAbc() As Boolean
        Get
            Dim dr As DataRow
            Dim i As Boolean = False

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                i = CBool(IIf(CInt(dr("abcplus_0")) = 2, True, False))
            End If

            Return i

        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("bpcplus_0") = IIf(value, 2, 1)
                dr.EndEdit()
            End If
        End Set
    End Property

    '2015-01-15 TABLA XA8 - 
    'Desarrollo de Agenda de Precios Clientes
    Public ReadOnly Property MostrarObsAgendaPrecio() As Boolean
        Get
            'Devuelve TRUE si la fecha actual es mayor o igual a la fecha
            'de la agenda-15 dias
            If dt2.Rows.Count > 0 Then
                Dim dr As DataRow = dt2.Rows(0)
                Dim f As Date = CDate(dr("fecha_0")).AddDays(-15)

                Return Date.Today >= f

            Else
                Return False

            End If
        End Get
    End Property
    Public ReadOnly Property ObsAgendaPrecio() As String
        Get
            Dim dr As DataRow = dt2.Rows(0)
            Return dr("obs_0").ToString.Trim
        End Get
    End Property

    ' IDisposable
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)

        If disposedValue Then Exit Sub

        ' TODO: Liberar otro estado (objetos administrados).
        If disposing Then

        End If

        ' TODO: Liberar su propio estado (objetos no administrados).
        ' TODO: Establecer campos grandes como Null.
        da.Dispose()
        dt.Dispose()

        MyBase.Dispose(disposing)

        Me.disposedValue = True

    End Sub

    'Function

End Class