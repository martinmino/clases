Imports System.Data.OracleClient
Imports System.IO

Public Class Pedido

    Private cn As OracleConnection
    Private WithEvents dah As OracleDataAdapter 'SORDER
    Private daq As OracleDataAdapter 'SORDERQ
    Private dap As OracleDataAdapter 'SORDERP
    Private dth As New DataTable
    Private dtq As New DataTable
    Private dtp As New DataTable
    Private ord As Orders

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String

        Me.cn = cn

        Sql = "SELECT * FROM sorder WHERE sohnum_0 = :sohnum_0"
        dah = New OracleDataAdapter(Sql, cn)
        dah.SelectCommand.Parameters.Add("sohnum_0", OracleType.VarChar)
        dah.InsertCommand = New OracleCommandBuilder(dah).GetInsertCommand
        dah.UpdateCommand = New OracleCommandBuilder(dah).GetUpdateCommand
        dah.DeleteCommand = New OracleCommandBuilder(dah).GetDeleteCommand

        Sql = "SELECT * FROM sorderq WHERE sohnum_0 = :sohnum_0"
        daq = New OracleDataAdapter(Sql, cn)
        daq.SelectCommand.Parameters.Add("sohnum_0", OracleType.VarChar)
        daq.InsertCommand = New OracleCommandBuilder(daq).GetInsertCommand
        daq.UpdateCommand = New OracleCommandBuilder(daq).GetUpdateCommand
        daq.DeleteCommand = New OracleCommandBuilder(daq).GetDeleteCommand

        Sql = "SELECT * FROM sorderp WHERE sohnum_0 = :sohnum_0"
        dap = New OracleDataAdapter(Sql, cn)
        dap.SelectCommand.Parameters.Add("sohnum_0", OracleType.VarChar)
        dap.InsertCommand = New OracleCommandBuilder(dap).GetInsertCommand
        dap.UpdateCommand = New OracleCommandBuilder(dap).GetUpdateCommand
        dap.DeleteCommand = New OracleCommandBuilder(dap).GetDeleteCommand

        dah.FillSchema(dth, SchemaType.Source)
        daq.FillSchema(dtq, SchemaType.Source)
        dap.FillSchema(dtp, SchemaType.Source)
    End Sub
    Public Sub Nuevo(ByVal ctz As Cotizacion)
        Dim dr As DataRow
        Dim bpc As Cliente = ctz.Cliente
        Dim bpa As Sucursal = ctz.Sucursal
        Dim cpy As Sociedad = ctz.Sociedad

        ord = New Orders(cn, ctz.Sociedad)

        dth.Clear()
        dtp.Clear()
        dtq.Clear()

        'SORDER
        dr = dth.NewRow
        dr("sohnum_0") = "0"
        dr("sohtyp_0") = "PED"
        dr("sohcat_0") = 1
        '*********************************************************
        '2016.10.04 - PEDIDO POR ISABEL
        'pedidos GRU de clientes que sean del vendedor 28
        'deben salir por planta G02
        '*********************************************************
        If bpc.Vendedor.Codigo = "28" AndAlso cpy.Codigo = "GRU" Then
            dr("salfcy_0") = "G02"
        Else
            dr("salfcy_0") = cpy.PlantaVenta(ctz.SaleDesdeMunro)
        End If
        dr("bpcord_0") = bpc.Codigo
        dr("bpcinv_0") = bpc.TerceroFactura
        dr("bpcpyr_0") = bpc.TerceroPagadorCodigo
        dr("bpcgru_0") = bpc.TerceroGrupoCodigo
        dr("bpaadd_0") = bpa.Sucursal
        dr("cusordref_0") = IIf(ctz.H, "H", ctz.OC).ToString
        dr("pjt_0") = " "
        dr("orddat_0") = Date.Today
        dr("vlydatcon_0") = #12/31/1599#
        dr("shidat_0") = Date.Today
        dr("demdlvdat_0") = Date.Today
        dr("lndrtndat_0") = #12/31/1599#
        dr("daylti_0") = 0
        dr("bpcnam_0") = bpc.Nombre
        dr("bpcnam_1") = " "
        dr("bpcaddlig_0") = bpa.Direccion
        dr("bpcaddlig_1") = " "
        dr("bpcaddlig_2") = " "
        dr("bpcposcod_0") = bpa.CodigoPostal
        dr("bpccty_0") = bpa.Ciudad
        dr("bpcsat_0") = bpa.Provincia
        dr("bpccry_0") = bpa.Pais.Codigo
        dr("bpccrynam_0") = bpa.Pais.Nombre
        dr("cntnam_0") = " " '--------------- VER COMO PONER PERSONA CONTACTO
        dr("bpinam_0") = bpc.Nombre
        dr("bpinam_1") = " "
        dr("bpiaddlig_0") = bpa.Direccion
        dr("bpiaddlig_1") = " "
        dr("bpiaddlig_2") = " "
        dr("bpiposcod_0") = bpa.CodigoPostal
        dr("bpicty_0") = bpa.Ciudad
        dr("bpisat_0") = bpa.Provincia
        dr("bpicry_0") = bpa.Pais.Codigo
        dr("bpicrynam_0") = bpa.Pais.Nombre
        dr("bpieecnum_0") = " "
        dr("cninam_0") = " " '--------------- VER COMO PONER PERSONA CONTACTO
        dr("bpdnam_0") = bpc.Nombre
        dr("bpdnam_1") = " "
        dr("bpdaddlig_0") = bpa.Direccion
        dr("bpdaddlig_1") = " "
        dr("bpdaddlig_2") = " "
        dr("bpdposcod_0") = bpa.CodigoPostal
        dr("bpdcty_0") = bpa.Ciudad
        dr("bpdsat_0") = bpa.Provincia
        dr("bpdcry_0") = bpa.Pais.Codigo
        dr("bpdcrynam_0") = bpa.Pais.Nombre
        dr("cndnam_0") = " " '--------------- VER COMO PONER PERSONA CONTACTO
        dr("rep_0") = bpc.Vendedor1Codigo
        dr("rep_1") = bpc.Vendedor2Codigo
        dr("rep_2") = bpc.Vendedor3Codigo
        dr("cur_0") = "ARS"
        dr("chgtyp_0") = 1
        dr("chgrat_0") = 1
        dr("cce_0") = " "
        dr("cce_1") = " "
        dr("cce_2") = " "
        dr("cce_3") = " "
        dr("cce_4") = " "
        dr("cce_5") = " "
        dr("cce_6") = " "
        dr("cce_7") = " "
        dr("cce_8") = " "
        dr("lan_0") = "SPA"
        dr("vacbpr_0") = bpc.RegimenImpuesto
        dr("pte_0") = ctz.CondicionPago
        dr("tsccod_0") = bpc.Familia1
        dr("tsccod_1") = bpc.Familia2
        dr("tsccod_2") = bpc.Familia3
        dr("tsccod_3") = bpc.Familia4
        dr("tsccod_4") = bpc.Familia5
        dr("dep_0") = " "
        dr("bptnum_0") = " "
        dr("mdl_0") = ctz.ModoEntrega
        dr("stofcy_0") = cpy.PlantaStock
        dr("drn_0") = bpa.SectorRuta
        dr("dlvpio_0") = bpa.PrioridadEntrega
        dr("ordcle_0") = 2
        dr("odl_0") = 1
        dr("dme_0") = 3
        dr("ime_0") = 6
        dr("ocnflg_0") = 1
        dr("ocnprn_0") = 1
        dr("sohtex1_0") = " "
        dr("sohtex2_0") = " "
        dr("sqhnum_0") = " "
        dr("betfcy_0") = 1
        dr("betcpy_0") = 1
        dr("orifcy_0") = " "
        dr("prfnum_0") = " "
        dr("lasdlvnum_0") = " "
        dr("lasdlvdat_0") = #12/31/1599#
        dr("lasinvnum_0") = " "
        dr("lasinvdat_0") = #12/31/1599#
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
        dr("prityp_0") = 1
        dr("ordnot_0") = 0 'Completar al recalcular lineas
        dr("ordati_0") = 0 'Completar al recalcular lineas
        dr("ordinvnot_0") = 0 'Completar al recalcular lineas
        dr("ordinvati_0") = 0 'Completar al recalcular lineas
        dr("dlrnot_0") = 0 'Completar al recalcular lineas
        dr("dlrati_0") = 0 'Completar al recalcular lineas
        dr("pfmtot_0") = 0 'Completar al recalcular lineas
        dr("discrgtyp_0") = 2
        dr("discrgtyp_1") = 2
        dr("discrgtyp_2") = 2
        dr("discrgtyp_3") = 2
        dr("discrgtyp_4") = 0
        dr("discrgtyp_5") = 0
        dr("discrgtyp_6") = 0
        dr("discrgtyp_7") = 0
        dr("discrgtyp_8") = 0
        dr("invdtalin_0") = 0
        dr("invdtalin_1") = 0
        dr("invdtalin_2") = 0
        dr("invdtalin_3") = 0
        dr("invdtalin_4") = 0
        dr("invdtalin_5") = 0
        dr("invdtalin_6") = 0
        dr("invdtalin_7") = 0
        dr("invdtalin_8") = 0
        dr("alltyp_0") = 1
        dr("unl_0") = 1
        dr("linnbr_0") = 0 'ACTUALIZAR: Numero de lineas del pedido
        dr("clelinnbr_0") = 0
        dr("alllinnbr_0") = 0 'ACTUALIZAR: Numero de lineas del pedido
        dr("dlvlinnbr_0") = 0
        dr("invlinnbr_0") = 0
        dr("ordsta_0") = 1
        dr("allsta_0") = 1
        dr("dlvsta_0") = 1
        dr("invsta_0") = 1
        dr("cdtsta_0") = 1
        dr("revnum_0") = 0
        dr("shiadecod_0") = " "
        dr("geocod_0") = " "
        dr("insctyflg_0") = " "
        dr("vtt_0") = " "
        dr("amttax_0") = 0
        dr("bastax_0") = 0
        dr("cclren_0") = " "
        dr("ccldat_0") = #12/31/1599#
        dr("srenum_0") = " "
        dr("cmgnum_0") = " "
        dr("opgnum_0") = " "
        dr("opgtyp_0") = " "
        dr("expnum_0") = 1
        dr("creusr_0") = USER
        dr("credat_0") = Date.Today
        dr("updusr_0") = " "
        dr("upddat_0") = #12/31/1599#
        dr("ysohdes_0") = ctz.Obs
        dr("yoriport_0") = " "
        dr("ydestport_0") = " "
        dr("ytripnum_0") = " "
        dr("xflgupd_0") = 0
        dr("yhdesde1_0") = ctz.HoraMananaDesde
        dr("yhhasta1_0") = ctz.HoraMananaHasta
        dr("yhdesde2_0") = ctz.HoraTardeDesde
        dr("yhhasta2_0") = ctz.HoraTardeHasta
        dr("ccn_0") = " "
        dr("xfcrto_0") = IIf(ctz.FcRto, 2, 1)
        dr("xweb_0") = " "
        dr("xgeo_0") = 2
        dr("ocf_0") = ctz.OCF 'Actualiza al subir la orden de compra 
        dr("tranum_0") = ctz.ExpresoCodigo
        dr("XFACTFLG_0") = 1
        dr("licitatyp_0") = ctz.TipoLicitacion
        dr("licitanum_0") = IIf(ctz.NumeroLicitacion = "", " ", ctz.NumeroLicitacion)
        dr("xsector_0") = EnvioAutomaticoSectorPedido(ctz)
        dr("xitnrecha_0") = ctz.IntervencionRechazo
        dr("xvtaantes_0") = IIf(ctz.AVentasAntesEntregar, 2, 1)
        dr("tcambio_0") = 0
        dr("demdlvdat_0") = IIf(ctz.FechaSolicitudEntrega < Today, Today, ctz.FechaSolicitudEntrega)

        dth.Rows.Add(dr)

        AgregarLineas(ctz)

    End Sub
    Public Sub Grabar()
        Dim dr As DataRow

        If dth.Rows.Count = 1 Then
            dr = dth.Rows(0)
            If dr.RowState = DataRowState.Added Then
                'La fila es nueva, asigno numero de pedido
                dr.BeginEdit()
                dr("sohnum_0") = AsignarNuevoNumero()
                dr.EndEdit()

                RecalcularLineas()

                ord.AgregarNecesidad(dtq)
                ord.Grabar()

                'Auto Asignacion
                AutoAsignacion()

                dah.Update(dth)
                daq.Update(dtq)
                dap.Update(dtp)

                'Comisiones
                Try
                    CargarComisiones()
                Catch ex As Exception
                End Try

            End If
            dah.Update(dth)
            daq.Update(dtq)
            dap.Update(dtp)
        End If
    End Sub
    Public Function Abrir(ByVal Nro As String) As Boolean
        dth.Clear()
        dtq.Clear()
        dtp.Clear()

        dah.SelectCommand.Parameters("sohnum_0").Value = Nro
        daq.SelectCommand.Parameters("sohnum_0").Value = Nro
        dap.SelectCommand.Parameters("sohnum_0").Value = Nro

        dah.Fill(dth)
        daq.Fill(dtq)
        dap.Fill(dtp)

        Return (dth.Rows.Count > 0)

    End Function
    Private Function AsignarNuevoNumero() As String
        Dim o As Numerador
        Dim n As Long 'Valor numerico
        Dim t As String 'Valor formateado (PED-D0211-0000)

        Dim n1 As String = "PED"
        Dim n2 As String = Planta.CodigoPlanta
        Dim n3 As String = Date.Today.ToString("yy")
        Dim n4 As String 'Numero

        'Periodo en 2 digitos
        o = New Numerador(cn, "PED", Planta.SociedadPlanta.Codigo, CInt(n3))
        n = o.Valor

        If n > 0 Then
            n4 = Strings.Right("00000" & n.ToString, 5)
            t = n1 & "-" & n2 & n3 & "-" & n4
        Else
            t = " "
        End If

        Return t

    End Function
    Public Sub AgregarLineas(ByVal ctz As Cotizacion)

        Dim dr As DataRow
        Dim itm As New Articulo(cn)

        For Each dr In ctz.Lineas.Rows
            If itm.Abrir(dr("itmref_0").ToString) Then
                'Si el articulo no es permitido para pedidos, salto al proximo registro
                If Not itm.ArticuloParaPedido Then Continue For

                AgregarLinea(itm, CDbl(dr("qty_0")), CDbl(dr("precio_0")), ctz)

            End If
        Next

    End Sub
    Private Sub AgregarLinea(ByVal itm As Articulo, ByVal Cantidad As Double, ByVal Precio As Double, ByVal ctz As Cotizacion)
        Dim dr As DataRow
        Dim linea As Long = LineaSiguiente()
        Dim vac As Impuesto = itm.Impuesto(ctz.Cliente)

        'SORDERQ - CANTIDADES
        dr = dtq.NewRow
        dr("sohnum_0") = Numero
        dr("soplin_0") = linea
        dr("soqseq_0") = linea
        dr("sohcat_0") = 1
        dr("salfcy_0") = ctz.Sociedad.PlantaVenta(ctz.SaleDesdeMunro)
        dr("bpcord_0") = ctz.ClienteCodigo
        dr("bpaadd_0") = ctz.SucursalCodigo
        dr("itmref_0") = itm.Codigo
        dr("stofcy_0") = ctz.Sociedad.PlantaStock
        dr("useplc_0") = " "
        dr("cad_0") = 0
        dr("yea_0") = 0
        dr("mon_0") = 0
        dr("dlvday_0") = 0
        dr("wee_0") = 0
        dr("orddat_0") = Date.Today
        dr("demdlvdat_0") = Date.Today
        dr("demdlvhou_0") = " "
        dr("demdlvref_0") = " "
        dr("impnumlig_0") = 0
        dr("shidat_0") = Date.Today
        dr("shihou_0") = " "
        dr("extdlvdat_0") = Date.Today
        dr("soqsta_0") = 1
        dr("invflg_0") = 1
        dr("demsta_0") = 1
        dr("demnum_0") = " " '!! VER QUE OCURRE SI NO SE LLENA !!
        dr("stomgtcod_0") = itm.GestionStock
        dr("lot_0") = " "
        dr("sta_0") = " "
        dr("loc_0") = " "
        dr("alltyp_0") = 1
        dr("oriqty_0") = Cantidad
        dr("qty_0") = Cantidad
        dr("shtqty_0") = 0
        dr("allqty_0") = 0
        dr("odlqty_0") = 0
        dr("dlvqty_0") = 0
        dr("invqty_0") = 0
        dr("tdlqty_0") = 0
        dr("qtystu_0") = Cantidad
        dr("shtqtystu_0") = 0
        dr("allqtystu_0") = 0
        dr("odlqtystu_0") = 0
        dr("dlvqtystu_0") = 0
        dr("invqtystu_0") = 0
        dr("tdlqtystu_0") = 0
        dr("drn_0") = ctz.Sucursal.Ruta
        dr("dlvpio_0") = ctz.Sucursal.PrioridadEntrega
        dr("dlvpiocmp_0") = 8 ' Lo pongo fijo pq no se que es
        dr("bptnum_0") = " "
        dr("mdl_0") = ctz.ModoEntrega
        dr("patnum_0") = " "
        dr("daylti_0") = 0
        dr("pck_0") = " "
        dr("pckcap_0") = 0
        dr("soqtex_0") = " "
        dr("sdhnum_0") = " "
        dr("sddlin_0") = 0
        dr("fmi_0") = 1
        dr("fminum_0") = " "
        dr("fmilin_0") = 0
        dr("fmiseq_0") = 0
        dr("pohnum_0") = " "
        dr("poplin_0") = 0
        dr("poqseq_0") = 0
        dr("perstrdat_0") = #12/31/1599#
        dr("perenddat_0") = #12/31/1599#
        dr("pernbrday_0") = 0
        dr("geocod_0") = " "
        dr("insctyflg_0") = " "
        dr("vts_0") = " "
        dr("vtc_0") = " "
        dr("taxgeoflg_0") = " "
        dr("taxflg_0") = 0
        dr("taxregflg_0") = 0
        dr("rattaxlin_0") = 0
        dr("bastaxlin_0") = 0
        dr("ocnprnbom_0") = 1
        dr("ndeprnbom_0") = 1
        dr("invprnbom_0") = 1
        dr("cclren_0") = " "
        dr("ccldat_0") = #12/31/1599#
        dr("pitflg_0") = 0
        dr("expnum_0") = 0
        dr("creusr_0") = USER
        dr("credat_0") = Date.Today
        dr("updusr_0") = " "
        dr("upddat_0") = #12/31/1599#
        dr("flgasig_0") = 0
        dr("tripnum_0") = " "
        dr("prep_0") = 0
        dtq.Rows.Add(dr)

        'SORDERP - PRECIOS
        dr = dtp.NewRow
        dr("sohnum_0") = Numero
        dr("soplin_0") = linea
        dr("sopseq_0") = linea
        dr("sohcat_0") = 1
        dr("strdat_0") = #12/31/1599#
        dr("enddat_0") = Date.Today
        dr("bpcord_0") = ctz.ClienteCodigo
        dr("bpaadd_0") = ctz.SucursalCodigo
        dr("cndnam_0") = " "
        dr("bpcinv_0") = ctz.Cliente.TerceroFactura
        dr("stofcy_0") = ctz.Sociedad.PlantaStock
        dr("salfcy_0") = ctz.Sociedad.PlantaVenta(ctz.SaleDesdeMunro)
        dr("itmref_0") = itm.Codigo
        dr("itmdes1_0") = itm.Descripcion
        dr("itmdes_0") = itm.Descripcion
        dr("itmrefbpc_0") = " " 'CHEQUEAR - es el codigo que el cliente usa para itm?
        dr("vacitm_0") = vac.Tipo
        dr("vacitm_1") = " "
        dr("vacitm_2") = " "
        dr("rep1_0") = ctz.Cliente.Vendedor1Codigo
        dr("rep2_0") = ctz.Cliente.Vendedor2Codigo
        dr("rep3_0") = ctz.Cliente.Vendedor3Codigo
        dr("reprat1_0") = 0
        dr("reprat2_0") = 0
        dr("reprat3_0") = 0
        dr("repcoe_0") = 1
        dr("gropri_0") = Precio
        dr("priren_0") = 2
        dr("netpri_0") = Precio
        dr("netprinot_0") = Precio
        If ctz.SociedadCodigo = "MON" Then
            dr("netpriati_0") = Precio
        Else
            dr("netpriati_0") = Precio + (Precio * vac.Alicuota / 100)
        End If
        dr("pfm_0") = Precio - itm.Costo
        dr("cprpri_0") = itm.Costo
        dr("discrgval1_0") = 0
        dr("discrgval2_0") = 0
        dr("discrgval3_0") = 0
        dr("discrgval4_0") = 0
        dr("discrgval5_0") = 0
        dr("discrgval6_0") = 0
        dr("discrgval7_0") = 0
        dr("discrgval8_0") = 0
        dr("discrgval9_0") = 0
        dr("discrgren1_0") = 0
        dr("discrgren2_0") = 0
        dr("discrgren3_0") = 0
        dr("discrgren4_0") = 0
        dr("discrgren5_0") = 0
        dr("discrgren6_0") = 0
        dr("discrgren7_0") = 0
        dr("discrgren8_0") = 0
        dr("discrgren9_0") = 0
        dr("vat_0") = vac.CodigoAlicuota
        dr("vat_1") = " "
        dr("vat_2") = " "
        dr("clcamt1_0") = 0
        dr("clcamt2_0") = 0
        dr("sau_0") = itm.UnidadVta
        dr("stu_0") = itm.Unidad
        dr("saustucoe_0") = 1
        dr("tsicod_0") = itm.Familia(0)
        dr("tsicod_1") = itm.Familia(1)
        dr("tsicod_2") = itm.Familia(2)
        dr("tsicod_3") = itm.Familia(3)
        dr("tsicod_4") = itm.Familia(4)
        dr("cce1_0") = itm.EjeAnalitico(0)
        dr("cce2_0") = itm.EjeAnalitico(1)
        dr("cce3_0") = itm.EjeAnalitico(2)
        dr("cce4_0") = itm.EjeAnalitico(3)
        dr("cce5_0") = itm.EjeAnalitico(4)
        dr("cce6_0") = itm.EjeAnalitico(5)
        dr("cce7_0") = itm.EjeAnalitico(6)
        dr("cce8_0") = itm.EjeAnalitico(7)
        dr("cce9_0") = itm.EjeAnalitico(8)
        dr("soqsta_0") = 1
        dr("lintyp_0") = 1
        dr("focflg_0") = 1
        dr("orilin_0") = 0
        dr("sqhnum_0") = " "
        dr("connum_0") = " "
        dr("sqdlin_0") = 0
        dr("linrevnum_0") = 0
        dr("expnum_0") = 1
        dr("creusr_0") = USER
        dr("credat_0") = Date.Today
        dr("updusr_0") = " "
        dr("upddat_0") = #12/31/1599#
        dr("xvat_0") = " "
        dr("xvat_1") = " "
        dr("xvat_2") = " "
        dr("xvat_3") = " "
        dr("xvat_4") = " "
        dr("xvat_5") = " "
        dr("xvat_6") = " "
        dr("xvat_7") = " "
        dr("xvat_8") = " "
        dr("xbasperpro_0") = 0
        dr("xtaxperpro_0") = 0
        dr("xtaxperpro_1") = 0
        dr("xtaxperpro_2") = 0
        dr("xtaxperpro_3") = 0
        dr("xtaxperpro_4") = 0
        dr("xtaxperpro_5") = 0
        dr("xtaxperpro_6") = 0
        dr("xtaxperpro_7") = 0
        dr("xtaxperpro_8") = 0
        dr("xporperpro_0") = 0
        dr("xporperpro_1") = 0
        dr("xporperpro_2") = 0
        dr("xporperpro_3") = 0
        dr("xporperpro_4") = 0
        dr("xporperpro_5") = 0
        dr("xporperpro_6") = 0
        dr("xporperpro_7") = 0
        dr("xporperpro_8") = 0
        dr("xvatint_0") = " "
        dr("xamttaxint_0") = 0
        dr("ygenitm_0") = itm.Codigo
        dr("xliqmrk_0") = 0
        dtp.Rows.Add(dr)

        RecalcularLineas()

    End Sub
    Private Sub RecalcularLineas()
        'Recorre todas las lineas del pedido y actualiza la cabecera
        Dim drq As DataRow
        Dim drp As DataRow
        Dim dr As DataRow
        Dim ordnot As Double = 0
        Dim ordati As Double = 0
        Dim ordinvnot As Double = 0
        Dim ordinvati As Double = 0
        Dim dlrnot As Double = 0
        Dim dlrati As Double = 0
        Dim pfmtot As Double = 0
        Dim i As Integer
        Dim NroPerido As String = " "

        Dim ord As New Orders(cn, Planta.SociedadPlanta)

        If dth.Rows.Count > 0 Then
            dr = dth.Rows(0)
            NroPerido = dr("sohnum_0").ToString
        End If

        For i = 0 To dtp.Rows.Count - 1
            drq = dtq.Rows(i)
            drp = dtp.Rows(i)

            'Actualizo numero de pedido en las lineas
            drp.BeginEdit()
            drp("sohnum_0") = NroPerido
            drp.EndEdit()

            drq.BeginEdit()
            drq("sohnum_0") = NroPerido
            drq.EndEdit()

            ordnot += CDbl(drp("netprinot_0")) * CDbl(drq("qty_0"))
            ordati += CDbl(drp("netpriati_0")) * CDbl(drq("qty_0"))
            ordinvnot += CDbl(drp("netprinot_0")) * CDbl(drq("qty_0"))
            ordinvati += CDbl(drp("netpriati_0")) * CDbl(drq("qty_0"))
            dlrnot += CDbl(drp("netprinot_0")) * CDbl(drq("qty_0"))
            dlrati += CDbl(drp("netpriati_0")) * CDbl(drq("qty_0"))
            pfmtot += CDbl(drp("pfm_0")) * CDbl(drq("qty_0"))

        Next

        'Actualizo tabla cabecera SORDER
        dr = dth.Rows(0)
        dr.BeginEdit()
        dr("ordnot_0") = ordnot
        dr("ordati_0") = ordati
        dr("ordinvnot_0") = ordinvnot
        dr("ordinvati_0") = ordinvati
        dr("dlrnot_0") = dlrnot
        dr("dlrati_0") = dlrati
        dr("pfmtot_0") = pfmtot
        dr("linnbr_0") = dtp.Rows.Count
        dr("alllinnbr_0") = dtp.Rows.Count
        dr.EndEdit()

    End Sub
    Private Function LineaSiguiente() As Long
        Dim i As Long = 0
        Dim dr As DataRow

        For Each dr In dtp.Rows
            If dr.RowState = DataRowState.Deleted Then Continue For
            If CLng(dr("soplin_0")) > i Then i = CLng(dr("soplin_0"))
        Next

        i += 1000

        Return i

    End Function
    Public Function ExportarArchivoMost(ByVal Archivo As String) As Integer
        Const SEP As String = "|"

        Dim bpc As Cliente = Me.Cliente
        Dim bpa As Sucursal = New Sucursal(cn, bpc.Codigo, Me.SucursalCodigo)
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim linea As String = ""
        Dim sw As StreamWriter '(Archivo, False, System.Text.Encoding.Default)
        Dim tab1 = New TablaVaria(cn, 22)
        Dim tab2 = New TablaVaria(cn, 21)
        Dim bomd As New Bomd(cn)

        'Valido que sucursal tenga cargado los datos del callejero
        If bpa.AlturaGCBA = 0 Or bpa.CallejeroGCBA = "" Then
            Return -1
        End If

        Sql = "SELECT zet.itmref_0, extract(year from zet.cardat_0) AS AA, itm.tsicod_2, itm.tsicod_1, zet.nrocil_0, mac.macnum_0, macpdtcod_0 "
        Sql &= "FROM (zetinue zet INNER JOIN machines mac ON (sernum_0 = macnum_0)) INNER JOIN itmmaster itm ON (zet.itmref_0 = itm.itmref_0) "
        Sql &= "WHERE sohnum_0 = :sohnum "
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("sohnum", OracleType.VarChar).Value = Me.Numero
        da.Fill(dt)

        sw = New StreamWriter(Archivo, False, System.Text.Encoding.Default)

        For Each dr In dt.Rows

            If bpc.TipoDoc = "80" And bpc.CUIT <> "" And bpc.Familia2 <> "11" Then
                'Empresa
                linea = "Empresa/Local" & SEP 'Campo 1
                linea &= bpc.CUIT & SEP 'Campo 2
                linea &= bpc.Nombre & SEP 'Campo 3
                linea &= SEP 'Campo 4
                linea &= SEP 'Campo 5
                linea &= SEP 'Campo 6

            Else
                'Particular
                linea = "Particular" & SEP 'Campo 1
                linea &= SEP 'Campo 2
                linea &= SEP 'Campo 3
                linea &= SEP 'Campo 4
                linea &= SEP 'Campo 5
                linea &= SEP 'Campo 6

            End If

            linea &= dr("nrocil_0").ToString & SEP 'Campo 7
            linea &= "False" & SEP 'Campo 8
            linea &= "Recarga" & SEP 'Campo 9
            linea &= tab1.Aux1(dr("tsicod_2").ToString) & SEP 'Campo 10
            linea &= tab2.Aux1(dr("tsicod_1").ToString) & SEP 'Campo 11
            linea &= "Matafuegos Donny S.R.L" & SEP 'Campo 12
            linea &= "Matafuegos Donny S.R.L" & SEP 'Campo 13
            linea &= SEP 'Campo 14
            linea &= dr("AA").ToString & SEP 'Campo 15

            bomd.Abrir(dr("macpdtcod_0").ToString, "20")
            linea &= Date.Today.AddDays(bomd.Dias).ToString("MM/yyyy") & SEP 'Campo 16

            linea &= bpa.CallejeroGCBA & SEP 'Campo 17
            linea &= bpa.AlturaGCBA & SEP 'Campo 18
            linea &= SEP 'Campo 19

            sw.WriteLine(linea)

        Next

        sw.Close()

        Return dt.Rows.Count

    End Function
    Public Sub CargarComisiones()
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim sql As String
        Dim dr As DataRow 'xvcrcom
        Dim drq As DataRow
        Dim drp As DataRow
        Dim comi As New Comision(cn)
        Dim itm As New Articulo(cn)

        If DB_USR <> "GEOPROD" Then Exit Sub

        sql = "SELECT * FROM xvcrcom WHERE vcrnum_0 = :vcrnum"
        da = New OracleDataAdapter(sql, cn)
        da.SelectCommand.Parameters.Add("vcrnum", OracleType.VarChar).Value = Me.Numero
        da.InsertCommand = New OracleCommandBuilder(da).GetInsertCommand
        da.DeleteCommand = New OracleCommandBuilder(da).GetDeleteCommand

        'Consulto registros existentes
        da.Fill(dt)

        'Elimino todos los registros encontrados
        For Each dr In dt.Rows
            dr.Delete()
        Next
        If dt.Rows.Count > 0 Then da.Update(dt)

        'Recorro todos los articulos del pedido
        For i = 0 To dtp.Rows.Count - 1
            drp = dtp.Rows(i)
            drq = dtq.Rows(i)

            itm.Abrir(drp("itmref_0").ToString)
            comi.Abrir(itm.Categoria, Cliente.CategoriaComision)

            Dim dr2 As DataRow = dt.NewRow
            dr2("vcrtyp_0") = 2
            dr2("vcrnum_0") = Me.Numero
            dr2("vcrlin_0") = drp("soplin_0")
            dr2("itmref_0") = itm.Codigo
            dr2("bprnum_0") = Cliente.Codigo
            dr2("tclcod_0") = itm.Categoria
            dr2("comcat_0") = Cliente.CategoriaComision
            dr2("rep_0") = Cliente.Vendedor.Codigo
            dr2("rep_1") = Cliente.Vendedor2.Codigo
            dr2("rep_2") = Cliente.Vendedor3.Codigo
            dr2("reprat_0") = comi.Comision(0)
            dr2("reprat_1") = comi.Comision(1)
            dr2("reprat_2") = comi.Comision(2)
            dr2("bprrat_0") = Cliente.Honorario
            dr2("bprrattyp_0") = 1
            dr2("cur_0") = "ARS"
            dr2("qty_0") = drq("qty_0")
            dr2("gropri_0") = drp("gropri_0")
            dr2("prinot_0") = drp("netprinot_0")
            dr2("priati_0") = drp("netpriati_0")
            dr2("stu_0") = drp("stu_0")
            dr2("orinum_0") = Me.Numero
            dr2("orityp_0") = 2
            dr2("orilin_0") = 0
            dr2("credat_0") = Date.Today
            dr2("creusr_0") = USER
            dr2("upddat_0") = #12/31/1599#
            dr2("updusr_0") = " "
            dr2("lotflg_0") = 0
            dr2("xflgupd_0") = 0
            dr2("xvcrlin_0") = 0

            dt.Rows.Add(dr2)

        Next

        If dt.Rows.Count > 0 Then da.Update(dt)

    End Sub

    Private Sub AutoAsignacion()
        'Se marca el pedido para auto asignacion se si cumple que
        'Sociedad: DNY
        'Clientes: 11
        'Todos los articulos tienen XAUTOASIG = 2
        Dim itm As New Articulo(cn)
        Dim flg As Boolean = False
        Dim dr As DataRow

        'Cliente debe ser 11 - Consorcio ó Consumidor Final
        'LA sociedad debe ser DNY
        If (Me.Cliente.Familia2 = "11" Or Me.Cliente.Codigo = "613613") And _
            Me.Planta.SociedadPlanta.Codigo = "DNY" Then flg = True

        'Recorro todos los articulos para ver si tienen auto asignacion
        For Each dr In dtq.Rows
            itm.Abrir(dr("itmref_0").ToString)
            If Not itm.AutoAsignacion Then flg = False
        Next
        'Si todos tienen autoasignacion, los marcos
        If flg Then
            'Marco cabecera del pedido
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("dlvpio_0") = IIf(flg, 2, 1)
            dr.EndEdit()
            'Marco detalle del pedido
            For Each dr In dtq.Rows
                dr.BeginEdit()
                dr("dlvpio_0") = 2
                dr.EndEdit()
            Next

        End If

    End Sub
    Private Function EnvioAutomaticoSectorPedido(ByVal ctz As Cotizacion) As String
        Dim CondicionPago As Integer
        Dim bpa As Sucursal
        Dim itn As New Intervencion(cn)
        Dim TieneRechazado As Boolean = False
        Dim Sector As String = ""

        bpa = ctz.Sucursal
        CondicionPago = CInt(ctz.CondicionPago)

        If ctz.AVentasAntesEntregar Then Return "VEN"

        If ctz.Cliente.Vendedor1Codigo = "28" Then Return "CTD"

        If ((CondicionPago >= 1 And CondicionPago <= 2 Or CondicionPago >= 10 And CondicionPago <= 23) And ctz.ExpresoCodigo = " ") Or _
           (CondicionPago >= 3 And CondicionPago <= 9) Or ctz.Sociedad.Codigo.ToString = "MON" Or ctz.ModoEntrega = "2" Then

            Return "ADM"
        End If

        If ((CondicionPago >= 1 And CondicionPago <= 2 Or CondicionPago >= 10 And CondicionPago <= 23) And ctz.ExpresoCodigo <> "" And bpa.Provincia = "CFE") Then
            Return "LOG"
        End If

        If ctz.IntervencionRechazo.Trim <> "" Then
            itn.Abrir(ctz.IntervencionRechazo)

            If itn.Sector = "CTD" Then
                Return "CTD"

            Else
                If itn.Tipo = "F1" Then
                    Return "ING"

                Else
                    If itn.Estado = 2 Or itn.Estado = 3 Or itn.Estado = 7 Then
                        Return "ADM"

                    Else
                        Return SectorDesconocido(ctz)

                    End If
                End If
            End If

        End If

        Return SectorDesconocido(ctz)

    End Function
    Private Function SectorDesconocido(ByVal ctz As Cotizacion) As String

        If ctz.TieneLineaDeProducto("401", "409", True) Or ctz.TieneLineaDeProducto("201", "202", False) And ctz.Sucursal.FechaInicioServicio > Date.Today.AddDays(-365) Then
            Return "ING"
        End If

        If ctz.TieneLineaDeProducto("651", "651", False) Or ctz.Cliente.EsAbonado = True Then
            Return "ABO"
        End If

        Return "LOG"

    End Function

    'PROPIEDADES
    Public ReadOnly Property Planta() As Planta
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return New Planta(cn, dr("salfcy_0").ToString)
        End Get
    End Property
    Public Property ModoEntrega() As String
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return dr("mdl_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dth.Rows(0)
            dr.BeginEdit()
            dr("mdl_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property Numero() As String
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return dr("sohnum_0").ToString
        End Get
    End Property
    Public ReadOnly Property Cliente() As Cliente
        Get
            Dim dr As DataRow = dth.Rows(0)
            Dim bpc As New Cliente(cn)
            bpc.Abrir(dr("bpcord_0").ToString)
            Return bpc
        End Get
    End Property
    Public ReadOnly Property SucursalCodigo() As String
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return dr("bpaadd_0").ToString
        End Get
    End Property
    Public ReadOnly Property Referencia() As String
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return dr("cusordref_0").ToString
        End Get
    End Property
    Public ReadOnly Property Fecha() As Date
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return CDate(dr("orddat_0"))
        End Get
    End Property
    Public Property Expreso() As String
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return dr("tranum_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dth.Rows(0)
            dr.BeginEdit()
            dr("tranum_0") = If(value.Trim = "", " ", value)
            dr.EndEdit()
        End Set
    End Property
    Public Property PresupuestoNumero() As String
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return dr("sqhnum_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dth.Rows(0)
            dr.BeginEdit()
            dr("sqhnum_0") = If(value.Trim = "", " ", value)
            dr.EndEdit()
        End Set
    End Property
    Public Property Sector() As String
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return dr("xsector_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dth.Rows(0)
            dr.BeginEdit()
            dr("xsector_0") = If(value.Trim = "", " ", value)
            dr.EndEdit()
        End Set
    End Property
    Public Property IntervencionRechazo() As String
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return dr("itnrecha_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dth.Rows(0)
            dr.BeginEdit()
            dr("xitnrecha_0") = If(value.Trim = "", " ", value)
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property VendedorCodigo() As String
        Get
            Dim dr As DataRow
            Dim v As String = " "

            If dth.Rows.Count > 0 Then
                dr = dth.Rows(0)
                v = dr("rep_0").ToString
            End If

            Return v
        End Get
    End Property
    Public Property TipoCambio() As Double
        Get
            Dim i As Double = 0

            If dth.Rows.Count > 0 Then
                Dim dr As DataRow = dth.Rows(0)
                i = CDbl(dr("tcambio_0"))
            End If

            Return i
        End Get
        Set(ByVal value As Double)

            If dth.Rows.Count = 0 Then Exit Property
            Dim dr As DataRow = dth.Rows(0)

            dr.BeginEdit()
            dr("tcambio_0") = value
            dr.EndEdit()
        End Set
    End Property

    Private Sub dah_RowUpdated(ByVal sender As Object, ByVal e As System.Data.OracleClient.OracleRowUpdatedEventArgs) Handles dah.RowUpdated
        If e.StatementType = StatementType.Insert Then
            Dim a As New Auditoria(cn)
            Dim dr As DataRow = e.Row

            a.Grabar(dr("creusr_0").ToString, "GESSOH", "SORDER", 1, dr("sohnum_0").ToString)
        End If
    End Sub
End Class