Imports System.Data.OracleClient

Public Class Presupuesto
    Const REPORT_NAME As String = "DEVTTC2.rpt"

    Private cn As OracleConnection
    Private da1 As OracleDataAdapter 'SQUOTE 
    Private da2 As OracleDataAdapter 'SQUOTED
    Private dt1 As New DataTable
    Private dt2 As New DataTable
    Private ds As New DataSet

    Public Sub New(ByVal cn As OracleConnection)
        Dim sql As String = ""
        Me.cn = cn

        sql = "SELECT * FROM squote WHERE sqhnum_0 = :sqhnum"
        da1 = New OracleDataAdapter(sql, cn)
        da1.SelectCommand.Parameters.Add("sqhnum", OracleType.VarChar)
        da1.UpdateCommand = New OracleCommandBuilder(da1).GetUpdateCommand
        da1.InsertCommand = New OracleCommandBuilder(da1).GetInsertCommand
        da1.DeleteCommand = New OracleCommandBuilder(da1).GetDeleteCommand

        sql = "SELECT * FROM squoted WHERE sqhnum_0 = :sqhnum"
        da2 = New OracleDataAdapter(sql, cn)
        da2.SelectCommand.Parameters.Add("sqhnum", OracleType.VarChar)
        da2.UpdateCommand = New OracleCommandBuilder(da2).GetUpdateCommand
        da2.InsertCommand = New OracleCommandBuilder(da2).GetInsertCommand
        da2.DeleteCommand = New OracleCommandBuilder(da2).GetDeleteCommand

        da1.FillSchema(dt1, SchemaType.Mapped)
        da2.FillSchema(dt2, SchemaType.Mapped)

        ds.Tables.Add(dt1)
        ds.Tables.Add(dt2)

        Dim fk As ForeignKeyConstraint
        fk = New ForeignKeyConstraint(dt1.Columns("sqhnum_0"), dt2.Columns("sqhnum_0"))
        fk.UpdateRule = Rule.Cascade
        fk.DeleteRule = Rule.Cascade

        dt2.Constraints.Add(fk)
    End Sub
    Public Sub Nuevo(ByVal ctz As Cotizacion)
        Dim dr As DataRow
        Dim bpc As Cliente = ctz.Cliente
        Dim bpa As Sucursal = ctz.Sucursal
        Dim cpy As Sociedad = ctz.Sociedad

        ds.Clear()

        dr = dt1.NewRow
        dr("sqhnum_0") = "0"
        dr("salfcy_0") = cpy.PlantaVenta
        dr("bpcord_0") = bpc.Cliente.Codigo
        dr("cusquoref_0") = " "
        dr("quodat_0") = Date.Today
        dr("vlydat_0") = Date.Today.AddDays(7)
        dr("daylti_0") = 0
        'Cliente pedido
        dr("bpaadd_0") = bpa.Sucursal
        dr("bpcnam_0") = bpc.Nombre
        dr("bpcnam_1") = " "
        dr("bpcaddlig_0") = bpc.SucursalDefault.Direccion
        dr("bpcaddlig_1") = " "
        dr("bpcaddlig_2") = " "
        dr("bpcposcod_0") = bpc.SucursalDefault.CodigoPostal
        dr("bpccty_0") = bpc.SucursalDefault.Ciudad
        dr("bpcsat_0") = bpc.SucursalDefault.Provincia
        dr("bpccry_0") = bpc.SucursalDefault.Pais.Codigo
        dr("bpccrynam_0") = bpc.SucursalDefault.Pais.Nombre
        dr("cncnam_0") = " "
        'Cliente Entregado
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
        dr("cndnam_0") = " "
        dr("rep_0") = bpc.Representante(0)
        dr("rep_1") = bpc.Representante(1)
        dr("rep_2") = bpc.Representante(2)
        dr("cur_0") = bpc.Divisa
        dr("chgtyp_0") = bpc.TipoCambio
        dr("chgrat_0") = 1
        dr("pjt_0") = " "
        dr("lan_0") = bpc.Idioma
        dr("vacbpr_0") = bpc.RegimenImpuesto
        dr("stofcy_0") = cpy.PlantaStock
        dr("prityp_0") = 1
        dr("quonot_0") = 0
        dr("quoati_0") = 0
        dr("quoinvnot_0") = 0
        dr("quoinvati_0") = 0
        dr("pfmtot_0") = 0
        dr("dep_0") = " "
        dr("pte_0") = ctz.CondicionPago
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
        dr("invdta_0") = "10"
        dr("invdta_1") = "0"
        dr("invdta_2") = "0"
        dr("invdta_3") = "0"
        dr("invdta_4") = "0"
        dr("invdta_5") = "0"
        dr("invdta_6") = "0"
        dr("invdta_7") = "0"
        dr("invdta_8") = "0"
        dr("invdta_9") = "0"
        dr("pbyprc_0") = "0"
        dr("sohnum_0") = " "
        dr("orddat_0") = #12/31/1599#
        dr("ordnbr_0") = 0
        dr("cfmlinnbr_0") = 0
        dr("linnbr_0") = 0 '¿cantidad de lineas?
        dr("quosta_0") = 1
        dr("quoprn_0") = 1
        dr("sqhtex1_0") = " "
        dr("sqhtex2_0") = " "
        dr("prfnum_0") = " "
        dr("geocod_0") = " "
        dr("insctyflg_0") = " "
        dr("vtt_0") = " "
        dr("amttax_0") = 0
        dr("bastax_0") = 0
        dr("expnum_0") = 1
        dr("creusr_0") = USER
        dr("credat_0") = Date.Today
        dr("updusr_0") = " "
        dr("upddat_0") = #12/31/1599#
        dr("xcoment_0") = ctz.Obs
        dr("x415_0") = 1
        dr("lstnum_0") = " "
        dr("xitn_0") = " "
        dr("dto_0") = 0
        dr("xratbpr_0") = bpc.Honorario
        dr("xsisfij_0") = 0
        dr("xdetec_0") = 1
        dr("xfm200_0") = 1
        dr("xco2_0") = 1
        dr("xflgren_0") = 1
        'Se agregaron los campos para los presupuestos de cocina
        dr("XVALDOL_0") = 0
        dr("XLTSACE_0") = 0
        dr("XTOBERAS_0") = 0
        dr("XFUSIBLES_0") = 0
        dr("XROLDANAS_0") = 0
        dr("XDIASINT_0") = 0
        dr("XDIASCAP_0") = 0
        dr("XMLCAMP_0") = 0
        dr("XANCAMP_0") = 0
        dr("XPROFCAMP_0") = 0
        dr("XDUCTOS_0") = 0
        dr("XOPERA_0") = 0
        dr("XVENT_0") = 0
        dr("XTOBERA1_0") = 0
        dr("XTOBERA2_0") = 0
        dr("XTOBERA3_0") = 0
        dr("XTOBERA4_0") = 0
        dr("XMTSCABLE_0") = 0
        dr("XCNTPAS_0") = 0
        dr("XCOSPAS_0") = 0
        dr("XCANTFLE_0") = 0
        dr("XFLEINT_0") = 0
        dr("VERELECSI_0") = 0
        dr("VERELECCI_0") = 0
        dr("VVERELECSI_0") = 0
        dr("VVERELECCI_0") = 0
        dr("XSELEGAS1_0") = 0
        dr("XSELEGAS12_0") = 0
        dr("XSELEGAS2_0") = 0
        dr("XEXTK_0") = 0
        dr("XDETGAS_0") = 0
        dr("XPANELTEMP_0") = 0
        dr("XMANO_0") = 0
        dr("XSISTEMA_0") = " "
        dr("XTIPOSIST_0") = 0

        dr("XAPROB_0") = 0
        dr("XUSRAPROB_0") = " "
        dr("XDATAPROB_0") = #12/31/1599#
        dr("XOPCION_0") = 0
        dr("XPRECIO_0") = 0
        dr("XAPROBMOTI_0") = 0

        dr("XVISITAS_0") = 0
        dr("XCOMPLEJO_0") = 0
        dr("XVALOR_0") = 0
        dr("XAEROSOL_0") = 0
        dr("XAEROSOLP_0") = 0
        dr("XDETECCION_0") = 0
        dr("XDETMARGEN_0") = 0
        dr("XDETDESC_0") = 0
        dr("licitatyp_0") = ctz.TipoLicitacion
        dr("licitanum_0") = IIf(ctz.NumeroLicitacion = "", " ", ctz.NumeroLicitacion)
        dr("tcambio_0") = 0
        dr("comi_0") = 0

        dt1.Rows.Add(dr)

        AgregarLineas(ctz)

    End Sub
    Public Function Abrir(ByVal Numero As String) As Boolean
        Dim flg As Boolean = False

        ds.Clear()
        da1.SelectCommand.Parameters("sqhnum").Value = Numero
        da2.SelectCommand.Parameters("sqhnum").Value = Numero

        Try
            da1.Fill(dt1)
            da2.Fill(dt2)

            flg = dt1.Rows.Count > 0

        Catch ex As Exception
            flg = False

        End Try

        Return flg

    End Function
    Public Sub Grabar()
        Dim dr As DataRow

        If dt1.Rows.Count = 1 Then
            dr = dt1.Rows(0)
            If dr.RowState = DataRowState.Added Then
                'La fila es nueva, asigno numero de pedido
                dr.BeginEdit()
                dr("sqhnum_0") = AsignarNuevoNumero()
                dr.EndEdit()

                RecalcularLineas()

            End If

            Try
                da1.Update(dt1)
                da2.Update(dt2)

            Catch ex As Exception
                Dim t As String = ex.Message
                t = t
            End Try

        End If

    End Sub
    Public Sub AgregarLineas(ByVal ctz As Cotizacion)

        Dim dr As DataRow
        Dim itm As New Articulo(cn)

        For Each dr In ctz.Lineas.Rows
            If itm.Abrir(dr("itmref_0").ToString) Then
                AgregarLinea(itm, CDbl(dr("qty_0")), CDbl(dr("precio_0")), ctz)
            End If
        Next

    End Sub
    Public Sub AprobarPresupuesto(ByVal usr As Usuario)
        Dim dr As DataRow

        If dt1.Rows.Count > 0 Then
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("XAPROB_0") = 2
            dr("XUSRAPROB_0") = usr.Codigo
            dr("XDATAPROB_0") = Date.Today
            dr.EndEdit()
        End If

    End Sub
    Private Sub AgregarLinea(ByVal itm As Articulo, ByVal Cantidad As Double, ByVal Precio As Double, ByVal ctz As Cotizacion)
        Dim dr As DataRow
        Dim linea As Long = LineaSiguiente()
        Dim vac As Impuesto = itm.Impuesto(ctz.Cliente)

        dr = dt2.NewRow
        dr("sqhnum_0") = CInt(dt1.Rows(0).Item("sqhnum_0"))
        dr("sqdlin_0") = linea
        dr("salfcy_0") = ctz.Sociedad.PlantaVenta
        dr("bpcord_0") = ctz.ClienteCodigo
        dr("quodat_0") = Date.Today
        dr("daylti_0") = "5"
        dr("itmref_0") = itm.Codigo
        dr("itmdes1_0") = itm.Descripcion
        dr("itmdes_0") = itm.Descripcion
        dr("vacitm_0") = "NOR"
        dr("vacitm_1") = " "
        dr("vacitm_2") = " "
        dr("rep1_0") = ctz.Cliente.Representante(0)
        dr("reprat1_0") = 0
        dr("rep2_0") = ctz.Cliente.Representante(1)
        dr("reprat2_0") = 0
        dr("rep3_0") = ctz.Cliente.Representante(2)
        dr("reprat3_0") = 0
        dr("repcoe_0") = 1
        dr("gropri_0") = Precio
        dr("priren_0") = 2
        dr("netpri_0") = Precio
        dr("pfm_0") = Precio - itm.Costo
        dr("netprinot_0") = Precio
        If ctz.SociedadCodigo = "MON" Then
            dr("netpriati_0") = Precio
        Else
            dr("netpriati_0") = Precio + (Precio * vac.Alicuota / 100)
        End If
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
        If ctz.Cliente.EsProspecto Then
            dr("vat_0") = "I21"
        Else
            dr("vat_0") = vac.CodigoAlicuota
        End If
        dr("vat_1") = " "
        dr("vat_2") = " "
        dr("clcamt1_0") = 0
        dr("clcamt2_0") = 0
        dr("qty_0") = Cantidad
        dr("sau_0") = itm.UnidadVta
        dr("stu_0") = itm.UnidadVta
        dr("saustucoe_0") = 1
        dr("bpaadd_0") = ctz.SucursalCodigo
        dr("stofcy_0") = ctz.Sociedad.PlantaStock
        dr("cndnam_0") = " "
        dr("lintyp_0") = 1
        dr("focflg_0") = 1
        dr("orilin_0") = 0
        dr("sohnum_0") = " "
        dr("soplin_0") = 0
        dr("ordflg_0") = 1
        dr("ordqty_0") = 0
        dr("sqdtex_0") = " "
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
        dr("expnum_0") = 1
        dr("creusr_0") = USER
        dr("credat_0") = Date.Today
        dr("updusr_0") = " "
        dr("upddat_0") = #12/31/1599#
        dt2.Rows.Add(dr)
    End Sub
    Private Function LineaSiguiente() As Long
        Dim i As Long = 0
        Dim dr As DataRow

        For Each dr In dt2.Rows
            If dr.RowState = DataRowState.Deleted Then Continue For
            If CLng(dr("sqdlin_0")) > i Then i = CLng(dr("sqdlin_0"))
        Next

        i += 1000

        Return i

    End Function
    Private Function AsignarNuevoNumero() As String
        Dim o As Numerador
        Dim n As Long 'Valor numerico
        Dim t As String 'Valor formateado (PRV-D0215-00001)

        Dim n1 As String = "PRV"
        Dim n2 As String = Planta.CodigoPlanta
        Dim n3 As String = Date.Today.ToString("yy")
        Dim n4 As String 'Numero

        'Periodo en 2 digitos
        o = New Numerador(cn, n1, Planta.SociedadPlanta.Codigo, CInt(n3))
        n = o.Valor

        If n > 0 Then

            n4 = Strings.Right("00000" & n.ToString, 5)
            t = n1 & "-" & n2 & n3 & "-" & n4
        Else
            t = " "
        End If

        Return t

    End Function
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
    Private Sub RecalcularLineas()
        'Recorre todas las lineas del presupuesto y actualiza la cabecera
        Dim drh As DataRow
        Dim drd As DataRow

        drh = dt1.Rows(0)
        drh.BeginEdit()
        drh("quonot_0") = 0
        drh("quoati_0") = 0
        drh("quoinvnot_0") = 0
        drh("quoinvati_0") = 0
        drh("pfmtot_0") = 0
        drh("linnbr_0") = dt2.Rows.Count
        drh.EndEdit()

        For i = 0 To dt2.Rows.Count - 1
            drd = dt2.Rows(i)

            drh.BeginEdit()
            'Actualizo numero de pedido en las lineas
            drh("quonot_0") = CDbl(drh("quonot_0")) + CDbl(drd("gropri_0")) * CDbl(drd("qty_0"))
            drh("quoati_0") = CDbl(drh("quoati_0")) + CDbl(drd("netpriati_0")) * CDbl(drd("qty_0"))
            drh("quoinvnot_0") = CDbl(drh("quoinvnot_0")) + CDbl(drd("gropri_0")) * CDbl(drd("qty_0"))
            drh("quoinvati_0") = CDbl(drh("quoinvati_0")) + CDbl(drd("netpriati_0")) * CDbl(drd("qty_0"))
            drh("pfmtot_0") = CDbl(drh("pfmtot_0")) + CDbl(drd("pfm_0")) * CDbl(drd("qty_0"))
            drh.EndEdit()

        Next

    End Sub
    Public Function PresupuestosPendientesAutorizacion(ByVal usr As Usuario, Optional ByVal dt As DataTable = Nothing) As DataTable
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim per As Permisos = usr.Permiso
        Dim EsAdmin As Boolean
        Dim b1, b2 As Boolean

        If dt Is Nothing Then dt = New DataTable

        EsAdmin = per.AccesoSecundario(73, "V")

        Sql = "select sqh.sqhnum_0, sqh.quodat_0, sqh.bpcord_0, sqh.bpaadd_0, sqh.bpcnam_0, sqh.bpdaddlig_0, bpc.rep_0, sqh.quoinvati_0, sqh.xaprob_0, sqh.xtiposist_0, bpa.xhidrante_0, sqh.xdeteccion_0, sqh.x415_0 "
        Sql &= "FROM squote sqh INNER JOIN "
        Sql &= "     bpcustomer bpc ON (sqh.bpcord_0 = bpc.bpcnum_0) INNER JOIN "
        Sql &= "     bpaddress  bpa ON (sqh.bpcord_0 = bpa.bpanum_0 AND sqh.bpaadd_0 = bpa.bpaadd_0) "
        Sql &= "WHERE sqh.xaprob_0 < 3 and "
        Sql &= "      (sqh.x415_0 = 2 or sqh.xdeteccion_0 = 2) and "
        Sql &= "      sqh.quodat_0 > to_date('15/03/2017', 'dd/mm/yyyy') and "
        Sql &= "      sqh.xaprob_0 <> 2 "
        Sql &= "ORDER BY xaprob_0 desc, quodat_0"

        da = New OracleDataAdapter(Sql, cn)
        dt.Clear()
        da.Fill(dt)

        If Not EsAdmin Then
            For Each dr As DataRow In dt.Rows
                Dim Rep As String = dr("rep_0").ToString
                b1 = per.TieneAccesoClienteAbonado(Rep)
                b2 = per.TieneAccesoClienteNoAbonado(Rep)

                If Not (b1 Or b2) Then dr.Delete()
            Next
        End If

        Return dt

    End Function
    Public Sub CrearLicitacion(ByVal Tipo As Integer, ByVal Numero As String)
        Me.TipoLicitacion = Tipo
        Me.NumeroLicitacion = Numero
        Me.Grabar()

        Dim lic As New Licitacion(cn)

        lic.Nueva()
        lic.setCliente(Me.Cliente)
        lic.setSucursal(Me.Sucursal)
        lic.TipoLicitacion = Me.TipoLicitacion
        lic.NumeroLicitacion = Me.NumeroLicitacion
        lic.NumeroPresupuesto = Me.Numero
        lic.FechaPresupuesto = Me.Fecha
        lic.TotalPresupuesto = Me.TotalII
        lic.Nuevo = TieneArticulosTipo(2)
        lic.Service = TieneArticulosTipo(1)
        lic.Agua = Me.Agua
        lic.Deteccion = Me.Deteccion

        'Aqui agregar detalle
        lic.AgregarDetalle(dt2)

        lic.Grabar()

    End Sub

    Public Function TieneArticulosTipo(ByVal Tipo As Integer) As Boolean
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable

        Sql = "SELECT xpe.* "
        Sql &= "FROM squoted sqd INNER JOIN xprecios xpe ON (sqd.itmref_0 = xpe.itmref_0) "
        Sql &= "WHERE ped_0 = :tipo AND sqhnum_0 = :sqhnum and xpe.itmref_0 <> 653001"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("sqhnum", OracleType.VarChar).Value = Numero
        da.SelectCommand.Parameters.Add("tipo", OracleType.Number).Value = Tipo

        da.Fill(dt)
        da.Dispose()

        Return dt.Rows.Count > 0

    End Function
    Public ReadOnly Property Numero() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("sqhnum_0").ToString
        End Get
    End Property
    Public ReadOnly Property Cliente() As Cliente
        Get
            Dim bpc As New Cliente(cn)
            Dim dr As DataRow = dt1.Rows(0)

            bpc.Abrir(dr("bpcord_0").ToString)

            Return bpc
        End Get
    End Property
    Public ReadOnly Property Sucursal() As Sucursal
        Get
            Dim bpa As New Sucursal(cn)
            Dim dr As DataRow = dt1.Rows(0)

            bpa.Abrir(dr("bpcord_0").ToString, dr("bpaadd_0").ToString)

            Return bpa
        End Get
    End Property
    Public ReadOnly Property Planta() As Planta
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return New Planta(cn, dr("salfcy_0").ToString)
        End Get
    End Property
    Public Property Vencimiento() As Date
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return CDate(dr("vlydat_0"))
        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow = dt1.Rows(0)
            dr.BeginEdit()
            dr("vlydat_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property Aprobado() As Boolean
        Get
            Dim dr As DataRow
            Dim x As Boolean = False

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                x = CBool(IIf(CInt(dr("XAPROB_0")) = 2, True, False))
            End If

            Return x

        End Get
    End Property
    Public ReadOnly Property AprobadoUsuario() As String
        Get
            Dim dr As DataRow
            Dim x As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                x = dr("XUSRAPROB_0").ToString
            End If

            Return x
        End Get
    End Property
    Public ReadOnly Property AprobadoFecha() As Date
        Get
            Dim dr As DataRow
            Dim x As Date = #12/31/1599#

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                x = CDate(dr("XDATAPROB_0"))
            End If

            Return x
        End Get
    End Property
    Public Property OpcionPago() As Integer
        Get
            Dim dr As DataRow
            Dim x As Integer

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                x = CInt(dr("XOPCION_0"))
            End If

            Return x
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("XOPCION_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property PrecioModificado() As Double
        Get
            Dim dr As DataRow
            Dim x As Double

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                x = CDbl(dr("XPRECIO_0"))
            End If

            Return x
        End Get
        Set(ByVal value As Double)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("XPRECIO_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public ReadOnly Property TotalAI() As Double
        Get
            Dim dr As DataRow
            Dim x As Double

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                x = CDbl(dr("quoinvnot_0"))
            End If

            Return x
        End Get

    End Property
    Public ReadOnly Property TotalII() As Double
        Get
            Dim dr As DataRow
            Dim x As Double

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                x = CDbl(dr("quoinvati_0"))
            End If

            Return x
        End Get

    End Property
    Public ReadOnly Property Fecha() As Date
        Get
            Dim dr As DataRow
            Dim x As Date = #12/31/1599#

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                x = CDate(dr("CREDAT_0"))
            End If

            Return x
        End Get
    End Property
    Public ReadOnly Property PrecioAprobacion() As Double
        Get
            If PrecioModificado > 0 Then
                Return PrecioModificado
            Else
                Return TotalII
            End If

        End Get
    End Property
    Public ReadOnly Property TipoSistema() As Integer
        Get
            Dim dr As DataRow
            Dim i As Integer = 0

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                i = CInt(dr("xtiposist_0"))
            End If

            Return i
        End Get
    End Property
    Public ReadOnly Property Renovacion() As Boolean
        Get
            Dim dr As DataRow
            Dim flg As Boolean = False

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                flg = CInt(dr("xflgren_0")) = 2
            End If

            Return flg
        End Get
    End Property
    Public Property EstadoAprobacion() As Integer
        Get
            Dim dr As DataRow
            Dim i As Integer = 0

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                i = CInt(dr("xaprob_0"))
            End If

            Return i
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("xaprob_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property MotivoPerdida() As Integer
        Get
            Dim dr As DataRow
            Dim i As Integer = 0

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                i = CInt(dr("xaprobmoti_0"))
            End If

            Return i
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("xaprobmoti_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property Comentarios() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("xcoment_0").ToString
            End If

            Return txt.Trim

        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            Dim txt As String

            If dt1.Rows.Count > 0 Then
                txt = value.Trim
                If txt = "" Then txt = " "

                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("xcoment_0") = txt
                dr.EndEdit()
            End If

        End Set
    End Property
    Public ReadOnly Property SucursalCodigo() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("bpaadd_0").ToString
            End If

            Return txt
        End Get
    End Property
    Public ReadOnly Property Agua() As Boolean
        Get
            Dim dr As DataRow
            Dim i As Boolean = False

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                i = CBool(IIf(dr("x415_0").ToString = "2", True, False))
            End If

            Return i
        End Get
    End Property
    Public ReadOnly Property Deteccion() As Boolean
        Get
            Dim dr As DataRow
            Dim i As Boolean = False

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                i = CBool(IIf(dr("xdeteccion_0").ToString = "2", True, False))
            End If

            Return i
        End Get
    End Property
    Public ReadOnly Property co2() As Boolean
        Get
            Dim dr As DataRow
            Dim i As Boolean = False

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                i = CBool(IIf(dr("xco2_0").ToString = "2", True, False))
            End If

            Return i
        End Get
    End Property
    Public ReadOnly Property tipodeteccion() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt2.Rows.Count > 0 Then
                dr = dt2.Rows(0)
                txt = dr("ITMREF_0").ToString
            End If

            Return txt
        End Get
    End Property
    Public ReadOnly Property fm200() As Boolean
        Get
            Dim dr As DataRow
            Dim i As Boolean = False

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                i = CBool(IIf(dr("xfm200_0").ToString = "2", True, False))
            End If

            Return i
        End Get
    End Property
    Public Property TipoLicitacion() As Integer
        Get
            Dim i As Integer = 0

            If dt1.Rows.Count > 0 Then
                Dim dr As DataRow = dt1.Rows(0)
                i = CInt(dr("licitatyp_0"))
            End If

            Return i
        End Get
        Set(ByVal value As Integer)

            If dt1.Rows.Count = 0 Then Exit Property
            Dim dr As DataRow = dt1.Rows(0)

            dr.BeginEdit()
            dr("licitatyp_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property NumeroLicitacion() As String
        Get
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                Dim dr As DataRow = dt1.Rows(0)
                txt = dr("licitanum_0").ToString
            End If

            Return txt.Trim
        End Get
        Set(ByVal value As String)
            If dt1.Rows.Count = 0 Then Exit Property
            Dim dr As DataRow = dt1.Rows(0)

            dr.BeginEdit()
            dr("licitanum_0") = IIf(value.Trim = "", " ", value)
            dr.EndEdit()
        End Set
    End Property
    Public Property TipoCambio() As Double
        Get
            Dim i As Double = 0

            If dt1.Rows.Count > 0 Then
                Dim dr As DataRow = dt1.Rows(0)
                i = CDbl(dr("tcambio_0"))
            End If

            Return i
        End Get
        Set(ByVal value As Double)

            If dt1.Rows.Count = 0 Then Exit Property
            Dim dr As DataRow = dt1.Rows(0)

            dr.BeginEdit()
            dr("tcambio_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Comision() As Double
        Get
            Dim i As Double = 0

            If dt1.Rows.Count > 0 Then
                Dim dr As DataRow = dt1.Rows(0)
                i = CDbl(dr("comi_0"))
            End If

            Return i
        End Get
        Set(ByVal value As Double)
            If dt1.Rows.Count = 0 Then Exit Property
            Dim dr As DataRow = dt1.Rows(0)

            dr.BeginEdit()
            dr("comi_0") = value
            dr.EndEdit()
        End Set
    End Property

End Class