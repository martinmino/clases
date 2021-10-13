Imports System.Data.OracleClient

Public Class ContratoServicio
    Public Event Mensaje(ByVal sender As Object, ByVal e As ContratoEventArgs)

    Private cn As OracleConnection
    Private da1 As OracleDataAdapter
    Private da2 As OracleDataAdapter
    Private da3 As OracleDataAdapter

    Private dt1 As New DataTable 'contserv
    Private dt2 As New DataTable 'contcob
    Private dt3 As New DataTable 'contsuc

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
        Adaptadores()

        da1.FillSchema(dt1, SchemaType.Mapped)
        da2.FillSchema(dt2, SchemaType.Mapped)
        da3.FillSchema(dt3, SchemaType.Mapped)

    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal Numero As String)
        Me.New(cn)
        Abrir(Numero)
    End Sub
    Private Sub Adaptadores()
        Dim Sql As String

        Sql = "SELECT * FROM contserv WHERE connum_0 = :connum_0"
        da1 = New OracleDataAdapter(Sql, cn)
        da1.SelectCommand.Parameters.Add("connum_0", OracleType.VarChar)

        da1.InsertCommand = New OracleCommandBuilder(da1).GetInsertCommand
        da1.UpdateCommand = New OracleCommandBuilder(da1).GetUpdateCommand
        da1.DeleteCommand = New OracleCommandBuilder(da1).GetDeleteCommand

        'Cobertura del contrato
        Sql = "SELECT cob.*, 0 AS qty_0, 0 AS rec FROM contcob cob WHERE connum_0 = :connum_0"
        da2 = New OracleDataAdapter(Sql, cn)
        da2.SelectCommand.Parameters.Add("connum_0", OracleType.VarChar)

        Sql = "INSERT INTO contcob VALUES (:connum_0, :covtyp_0, :cod_0, :covqty_0)"
        da2.InsertCommand = New OracleCommand(Sql, cn)

        Sql = "UPDATE contcob SET connum_0 = :connum_0, covtyp_0 = :covtyp_0, cod_0 = :cod_0, covqty_0 = :covqty_0 "
        Sql &= "WHERE connum_0 = :connum_0w, covtyp_0 = :covtyp_0w, cod_0 = :cod_0w "
        da2.UpdateCommand = New OracleCommand(Sql, cn)

        Sql = "DELETE FROM contcob WHERE connum_0 = :connum_0 AND covtyp_0 = :covtyp_0 AND cod_0 = :cod_0"
        da2.DeleteCommand = New OracleCommand(Sql, cn)


        With da2
            .SelectCommand.Parameters.Add("connum_0", OracleType.VarChar)

            Parametro(.InsertCommand, "connum_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "covtyp_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "cod_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "covqty_0", OracleType.Number, DataRowVersion.Current)

            Parametro(.UpdateCommand, "connum_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "covtyp_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "cod_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "covqty_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "connum_0w", OracleType.VarChar, DataRowVersion.Original, "connum_0")
            Parametro(.UpdateCommand, "covtyp_0w", OracleType.Number, DataRowVersion.Original, "covtyp_0")
            Parametro(.UpdateCommand, "cod_0w", OracleType.VarChar, DataRowVersion.Original, "cod_0")

            Parametro(.DeleteCommand, "connum_0", OracleType.VarChar, DataRowVersion.Original)
            Parametro(.DeleteCommand, "covtyp_0", OracleType.Number, DataRowVersion.Original)
            Parametro(.DeleteCommand, "cod_0", OracleType.VarChar, DataRowVersion.Original)

        End With

        Sql = "SELECT * FROM contsuc WHERE connum_0 = :connum_0"
        da3 = New OracleDataAdapter(Sql, cn)
        da3.SelectCommand.Parameters.Add("connum_0", OracleType.VarChar)
        da3.InsertCommand = New OracleCommandBuilder(da3).GetInsertCommand
        da3.UpdateCommand = New OracleCommandBuilder(da3).GetUpdateCommand
        da3.DeleteCommand = New OracleCommandBuilder(da3).GetDeleteCommand

    End Sub
    Public Sub Nuevo(ByVal bpc As Cliente)
        Dim dr As DataRow

        dt1.Clear()
        dt2.Clear()
        dt3.Clear()

        dr = dt1.NewRow
        dr("connum_0") = "0"
        dr("connam_0") = " "
        dr("cpy_0") = "DNY"
        dr("salfcy_0") = "D02"
        dr("conbpc_0") = bpc.Codigo
        dr("conbpcinv_0") = bpc.TerceroFactura
        dr("conbpcpyr_0") = bpc.TerceroPagador.Codigo
        dr("conbpcgru_0") = bpc.TerceroGrupoCodigo
        dr("bpaadd_0") = bpc.SucursalDefault.Sucursal
        dr("yref_0") = " "
        dr("conccn_0") = " "
        dr("conpjt_0") = " "
        For i = 0 To 4
            dr("contyp_" & i.ToString) = " "
        Next
        dr("contypcla_0") = " "
        dr("gua_0") = 0
        dr("concat_0") = 0
        dr("siudat_0") = #12/31/1599#
        dr("constrdat_0") = #12/31/1599#
        dr("oconstrdat_0") = #12/31/1599#
        dr("conenddat_0") = #12/31/1599#
        dr("oconenddat_0") = #12/31/1599#
        dr("rewfry_0") = 0
        dr("rewfrybas_0") = 4
        dr("invfry_0") = 0
        dr("invfrybas_0") = 4
        dr("invmet_0") = 0
        dr("invfrycoe_0") = 0
        dr("invlti_0") = 7
        dr("invltibas_0") = 2
        dr("lasinvdat_0") = #12/31/1599#
        dr("olasinvdat_0") = #12/31/1599#
        dr("onexinvdat_0") = #12/31/1599#
        dr("nexinvdat_0") = #12/31/1599#
        dr("onexinvamt_0") = 0
        dr("nexinvamt_0") = 0
        dr("onexshiinv_0") = #12/31/1599#
        dr("nexshiinv_0") = #12/31/1599#
        dr("pte_0") = "001"
        dr("conchgtyp_0") = 0
        dr("conprityp_0") = 0
        dr("convacbpr_0") = bpc.RegimenImpuesto
        dr("itmref_0") = " "
        dr("rvafry_0") = 0
        dr("rvafrybas_0") = 0
        dr("rvamet_0") = 0
        dr("rvassp_0") = 0
        dr("rvadat_0") = #12/31/1599#
        dr("orvadat_0") = #12/31/1599#
        dr("conamt_0") = 0
        dr("oconamt_0") = 0
        dr("cur_0") = " "
        dr("pitcdt_0") = 0
        dr("pitrer_0") = 0
        dr("pitcsm_0") = 0
        dr("pitblc_0") = 0
        dr("pittol_0") = 0
        dr("salrep_0") = bpc.Vendedor1Codigo
        dr("salrep_1") = bpc.Vendedor2Codigo
        dr("salrep_2") = bpc.Vendedor3Codigo
        dr("conbasidx_0") = " "
        dr("conbasfor_0") = " "
        dr("lasvalidx_0") = 0
        dr("olasvalidx_0") = 0
        dr("lasidxdat_0") = #12/31/1599#
        dr("olasidxdat_0") = #12/31/1599#
        dr("conrew_0") = 0
        dr("manrewflg_0") = 0
        dr("manrewamt_0") = 0
        dr("rsilti_0") = 0
        dr("rsiltibas_0") = 0
        dr("evrpbl_0") = 0
        dr("evrmac_0") = 0
        dr("crscovsam_0") = 0
        dr("ittfcy_0") = 0
        dr("ittdur_0") = 0
        dr("ittdurbas_0") = 0
        dr("ittdatend_0") = #12/31/1599#
        dr("ittltimax_0") = 0
        dr("basexxitt_0") = 0
        dr("solltimax_0") = 0
        dr("basexxsol_0") = 0
        dr("amtply_0") = 0
        dr("basexxamt_0") = 0
        dr("othcni_0") = " "
        dr("ittfcyx_0") = 0
        dr("ittdurx_0") = 0
        dr("ittdurbasx_0") = 0
        dr("ittdatendx_0") = #12/31/1599#
        dr("ittltimaxx_0") = 0
        dr("basexxittx_0") = 0
        dr("solltimaxx_0") = 0
        dr("basexxsolx_0") = 0
        dr("amtplyx_0") = 0
        dr("basexxamtx_0") = 0
        dr("rsiflg_0") = 0
        dr("rsidat_0") = #12/31/1599#
        dr("rsiren_0") = 0
        dr("fddflg_0") = 0
        dr("fddusr_0") = " "
        dr("ordnum_0") = " "
        dr("ordlinnum_0") = 0
        dr("ordupdflg_0") = 0
        For i = 0 To 24
            dr("acpdptcod_" & i.ToString) = " "
        Next
        dr("oinv_0") = " "
        dr("orewinv_0") = " "
        dr("orvainv_0") = " "
        For i = 0 To 9
            dr("invdtaamt_" & i.ToString) = 0
            dr("invdta_" & i.ToString) = 0
        Next
        For i = 0 To 8
            dr("cce_" & i.ToString) = " "
        Next
        dr("conori_0") = 0
        dr("conoritxt_0") = " "
        dr("conorivcr_0") = 0
        dr("conorivcrl_0") = 0
        dr("concot_0") = 0
        dr("dsywnd_0") = 1
        dr("odsywnd_0") = 0
        dr("creusr_0") = " "
        dr("credat_0") = Date.Today
        dr("updusr_0") = " "
        dr("upddat_0") = #12/31/1599#
        dr("ypctpres_0") = 0
        dr("xsuspend_0") = 1
        dr("xunifi_0") = 0

        dt1.Rows.Add(dr)

    End Sub
    Public Function Abrir(ByVal Nro As String) As Boolean
        Dim dr As DataRow
        Dim da As OracleDataAdapter = Nothing
        Dim Sql As String
        Dim dt As New DataTable

        dt1.Clear()
        da1.SelectCommand.Parameters("connum_0").Value = Nro
        da1.Fill(dt1)

        'Cobertura del contrato
        dt2.Clear()
        da2.SelectCommand.Parameters("connum_0").Value = Nro
        da2.Fill(dt2)

        'Sucursales cubiertas
        dt3.Clear()
        da3.SelectCommand.Parameters("connum_0").Value = Nro
        da3.Fill(dt3)

        'Calculo el uso de la cobertura
        For Each dr In dt2.Rows
            dt.Clear()

            Select Case CInt(dr("covtyp_0"))
                Case 2
                    Sql = "SELECT SUM(hdk.hdtqty_0) "
                    Sql &= "FROM serrequest sre INNER JOIN interven itn ON (sre.srenum_0 = itn.srvdemnum_0) "
                    Sql &= "     INNER JOIN hdktask hdk ON (hdk.srenum_0 = sre.srenum_0 AND hdk.itnnum_0 = itn.num_0) "
                    Sql &= "WHERE sre.conspt_0 = :connum_0 AND don_0 = 2 AND itn.zflgtrip_0 < 8 AND ymrkitn_0 = 2 AND hdtitm_0 = :hdtitm_0"

                    da = New OracleDataAdapter(Sql, cn)
                    da.SelectCommand.Parameters.Add("connum_0", OracleType.VarChar).Value = Nro
                    da.SelectCommand.Parameters.Add("hdtitm_0", OracleType.VarChar).Value = dr("cod_0").ToString

                Case 4
                    Sql = "SELECT SUM(hdk.hdtqty_0)"
                    Sql &= "FROM serrequest sre INNER JOIN interven itn ON (sre.srenum_0 = itn.srvdemnum_0)"
                    Sql &= "	 INNER JOIN hdktask hdk ON (hdk.srenum_0 = sre.srenum_0 AND hdk.itnnum_0 = itn.num_0)"
                    Sql &= "	 INNER JOIN itmmaster itm ON (hdk.hdtitm_0 = itm.itmref_0)"
                    Sql &= "WHERE conspt_0 = :connum_0 AND don_0 = 2 AND itn.zflgtrip_0 < 8 AND ymrkitn_0 = 2 AND itm.cfglin_0 = :cfglin_0"

                    da = New OracleDataAdapter(Sql, cn)
                    da.SelectCommand.Parameters.Add("connum_0", OracleType.VarChar).Value = Nro
                    da.SelectCommand.Parameters.Add("cfglin_0", OracleType.VarChar).Value = dr("cod_0").ToString

            End Select

            da.Fill(dt)

            dr.BeginEdit()
            If dt.Rows.Count = 0 OrElse IsDBNull(dt.Rows(0).Item(0)) Then
                dr("qty_0") = 0
            Else
                dr("qty_0") = CInt(dt.Rows(0).Item(0))
            End If
            dr.EndEdit()

        Next

        Return dt1.Rows.Count > 0

    End Function
    Public Function Grabar() As Boolean

        'Se asigna numero al contrato
        If Numero = "0" Then
            'Obtengo el próximo número a asignar
            Dim n As String = ProximoNumero()
            Dim dr As DataRow

            For Each dr In dt1.Rows
                dr.BeginEdit()
                dr("connum_0") = n
                dr.EndEdit()
            Next
            For Each dr In dt2.Rows
                dr.BeginEdit()
                dr("connum_0") = n
                dr.EndEdit()
            Next
            For Each dr In dt3.Rows
                dr.BeginEdit()
                dr("connum_0") = n
                dr.EndEdit()
            Next

        End If

        da1.Update(dt1)
        da2.Update(dt2)
        da3.Update(dt3)

    End Function
    Private Function ProximoNumero() As String
        Dim o As Numerador
        Dim n As Long 'Valor numerico
        Dim t As String 'Valor formateado (PRV-D0215-00001)

        Dim n1 As String = Date.Today.ToString("yy")
        Dim n2 As String 'Numero

        o = New Numerador(cn, "CON")

        n = o.Valor

        If n > 0 Then
            n2 = Strings.Right("00000" & n.ToString, 6)
            t = n1 & n2

        Else
            t = " "

        End If

        Return t

    End Function
    Public Sub AgregarCobertura(ByVal Tipo As Integer, ByVal Codigo As String, ByVal Cantidad As Integer)
        Dim dr As DataRow = Nothing
        Dim existe As Boolean = False

        'Busco si el codigo existe
        For Each dr In dt2.Rows
            If dr.RowState = DataRowState.Deleted Then Continue For

            If dr("cod_0").ToString = Codigo Then
                existe = True
                Exit For
            End If
        Next

        If Not existe Then
            dr = dt2.NewRow
            dr("connum_0") = Numero
            dr("covtyp_0") = Tipo
            dr("cod_0") = Codigo
            dr("covqty_0") = Cantidad
            dt2.Rows.Add(dr)

        Else
            dr.BeginEdit()
            dr("connum_0") = Numero
            dr("covtyp_0") = Tipo
            dr("cod_0") = Codigo
            dr("covqty_0") = Cantidad
            dr.EndEdit()
        End If

    End Sub
    Public Sub AgregarSucursalCubierta(ByVal CodigoSucursal As String)
        Dim dr As DataRow

        dr = dt3.NewRow
        dr("connum_0") = Numero
        dr("bpaadd_0") = CodigoSucursal
        dr("bpcnum_0") = Me.Cliente
        dt3.Rows.Add(dr)
    End Sub
    Public Sub AgregarSucursales(ByVal dt As DataTable)
        'Limpio todas las sucursales
        For Each dr As DataRow In dt3.Rows
            dr.Delete()
        Next
        'Agrego las nuevas sucursales
        For Each dr As DataRow In dt.Rows
            AgregarSucursalCubierta(dr("bpaadd_0").ToString)
        Next
    End Sub
    Public Function IncluyeSucursal(ByVal Suc As String) As Boolean
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim flg As Boolean = False

        For Each dr In dt3.Rows
            If dr("bpaadd_0").ToString = Suc Then
                flg = True
                Exit For
            End If
        Next

        Return flg

    End Function
    Public Function ArticuloCubierto(ByVal Art As String) As Boolean
        Dim itm As New Articulo(cn)
        Dim flg As Boolean = False

        If Not itm.Abrir(Art) Then Exit Function

        For Each dr As DataRow In dt2.Rows
            If dr("cod_0").ToString = Art Or dr("cod_0").ToString = itm.LineaProducto Then
                flg = True
                Exit For
            End If
        Next

        itm.Dispose()

        Return flg

    End Function
    Private Function ArticuloCubierto(ByVal Codigo As String, ByVal Cant As Integer) As Integer
        Dim dr As DataRow
        Dim i As Integer = -1

        For Each dr In dt2.Rows

            If CInt(dr("covtyp_0")) = 2 AndAlso dr("cod_0").ToString = Codigo Then
                If CInt(dr("covqty_0")) >= Cant Then
                    i = 0

                Else
                    i = Cant - CInt(dr("covqty_0"))

                End If

                Exit For

            End If
        Next

        Return i

    End Function
    Public Function VerificarCobertura(ByVal dt As DataTable, ByVal bpa As String) As Boolean
        Dim dr As DataRow
        Dim itm As New Articulo(cn)
        Dim flg As Boolean = False
        Dim e As New ContratoEventArgs
        Dim i As Integer

        'Reset acmuladores
        For Each dr In dt2.Rows
            dr.BeginEdit()
            dr("rec") = dr("qty_0")
            dr.EndEdit()
        Next

        'Recorro todos los servicios cargado en la intervencion
        For Each dr In dt.Rows
            If CInt(dr("typlig_0")) <> 1 Then Continue For

            itm.Abrir(dr("itmref_0").ToString)

            'Evaluacion de estado de cobertura
            i = ArticuloCubierto(dr("itmref_0").ToString, CInt(dr("tqty_0")))

            Select Case i
                Case -1 'Articulo no incluido
                    'Compruebo si la linea de producto del articulo esta cubierta
                    If LineaCubierta(itm.LineaProducto, CInt(dr("tqty_0"))) Then
                        flg = True
                    Else
                        e.Mensaje = "Articulo " & itm.Codigo & " no cubierto por el contrato"
                        flg = False
                        Exit For
                    End If

                Case 0 'Articulo incluido y no sobrepasado
                    LineaCubierta(itm.LineaProducto, CInt(dr("tqty_0")))
                    flg = True

                Case Else 'Articulo incluido y cantidad superada
                    e.Mensaje = itm.Codigo & " - Se sobrepasa la cobertura para este artículo." & vbCrLf
                    e.Mensaje &= "Cantidad excedente: " & i.ToString
                    flg = False
                    Exit For

            End Select

        Next

        If flg Then
            'Por cada linea de producto chequeo que no se supere la cantidad cubierta
            For Each dr In dt2.Rows
                Dim Exceso As Integer

                Exceso = CInt(dr("covqty_0")) - CInt(dr("rec"))
                Exceso = Math.Abs(Exceso)

                If CInt(dr("covtyp_0")) = 4 AndAlso CInt(dr("covqty_0")) < CInt(dr("rec")) Then
                    e.Mensaje = "Se sobrepasa la cobertura en " & Exceso & " para la línea de producto " & dr("cod_0").ToString & " en el contrato " & Me.Numero & " del cliente " & Me.Cliente & "-" & bpa & " - Unificacion: " & Me.Unificacion.ToString & " meses"

                    flg = False
                    Exit For
                End If
            Next
        End If

        If Not flg Then RaiseEvent Mensaje(Me, e)

        Return flg

    End Function
    Public Sub EliminarCoberturaExcedida(ByVal dt As DataTable, ByVal bpa As Sucursal)
        Dim dr As DataRow
        Dim itm As New Articulo(cn)
        Dim flg As Boolean = False
        Dim e As New ContratoEventArgs
        Dim i As Integer
        Dim LineasArr As New ArrayList

        'Reset acmuladores
        For Each dr In dt2.Rows
            dr.BeginEdit()
            dr("rec") = dr("qty_0")
            dr.EndEdit()
        Next

        'Recorro todos los servicios cargado en la intervencion
        For j As Integer = dt.Rows.Count - 1 To 0 Step -1
            dr = dt.Rows(j)

            If CInt(dr("typlig_0")) <> 1 Then Continue For

            itm.Abrir(dr("itmref_0").ToString)

            'Evaluacion de estado de cobertura
            i = ArticuloCubierto(dr("itmref_0").ToString, CInt(dr("tqty_0")))

            flg = True

            Select Case i
                Case -1 'Articulo no incluido
                    'Compruebo si la linea de producto del articulo esta cubierta
                    If LineaCubierta(itm.LineaProducto, CInt(dr("tqty_0"))) Then
                        flg = True
                    Else
                        e.Mensaje = "Articulo " & itm.Codigo & " no cubierto por el contrato " & Me.Numero
                        flg = False

                        dr.Delete()
                    End If

                Case 0 'Articulo incluido y no sobrepasado
                    LineaCubierta(itm.LineaProducto, CInt(dr("tqty_0")))
                    flg = True

                Case Else 'Articulo incluido y cantidad superada

                    e.Mensaje = "Se supera la cobertura en " & i.ToString & " para el artículo " & itm.Codigo & " en el contrato " & Me.Numero & " del cliente " & Me.Cliente

                    flg = False

                    dr.Delete()

            End Select

            If Not flg Then RaiseEvent Mensaje(Me, e)

        Next

        'Por cada linea de producto chequeo que no se supere la cantidad cubierta
        For Each dr In dt2.Rows
            If dr.RowState = DataRowState.Deleted Then Continue For

            If CInt(dr("covtyp_0")) = 4 AndAlso CInt(dr("covqty_0")) < CInt(dr("rec")) Then
                e.Mensaje = "Se sobrepasa la cobertura para la línea de producto " & dr("cod_0").ToString & " en el contrato " & Me.Numero & " del cliente " & Me.Cliente
                'e.Mensaje &= "Cobertura: " & dr("covqty_0").ToString & " - Consumido: " & dr("rec").ToString

                'Agrego al Array la linea de producto excedida
                LineasArr.Add(dr("cod_0").ToString)

            End If
        Next

        'Recorro todas las lineas de producto excedidas y borro los articulos incluidos en esas lineas
        For i = 0 To LineasArr.Count - 1
            Dim l As String = LineasArr.Item(i).ToString
            For j = dt2.Rows.Count - 1 To 0 Step -1
                dr = dt2.Rows(j)

                If dr.RowState = DataRowState.Deleted Then Continue For
                If CInt(dr("covtyp_0")) <> 2 Then Continue For

                itm.Abrir(dr("cod_0").ToString)
                If itm.LineaProducto = l Then dr.Delete()

            Next
        Next

        dt.AcceptChanges()

    End Sub
    Public Function CoberturaPorArticulo(ByVal Codigo As String) As Integer
        Dim dr As DataRow
        Dim i As Integer = 0

        For Each dr In dt2.Rows
            If dr("cod_0").ToString = Codigo Then
                i = CInt(dr("covqty_0"))
                Exit For
            End If
        Next

        Return i

    End Function
    Public Function CantidadCubierta(ByVal Codigo As String) As Integer
        'Devuelve la cantidad cubierta del codigo consultado
        Dim i As Integer = 0
        Dim dr As DataRow

        For Each dr In dt2.Rows
            If dr("cod_0").ToString = Codigo Then
                i += CInt(dr("covqty_0"))
            End If
        Next

        Return i
    End Function
    Private Function LineaCubierta(ByVal Linea As String, ByVal cant As Integer) As Boolean
        Dim dr As DataRow
        Dim flg As Boolean = False

        For Each dr In dt2.Rows
            If CInt(dr("covtyp_0")) = 4 AndAlso dr("cod_0").ToString = Linea Then
                dr.BeginEdit()
                dr("rec") = CInt(dr("rec")) + cant
                dr.EndEdit()

                flg = True

                Exit For

            End If
        Next

        Return flg

    End Function


    Public ReadOnly Property Rescindido() As Boolean
        Get
            Dim dr As DataRow
            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                Return CInt(dr("rsiflg_0")) = 2
            Else
                Return True
            End If
        End Get
    End Property
    Public ReadOnly Property Cerrado() As Boolean
        Get
            Dim dr As DataRow
            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                Return CInt(dr("fddflg_0")) = 2
            Else
                Return True
            End If
        End Get
    End Property
    Public ReadOnly Property ContratoVigente(ByVal Fecha As Date) As Boolean
        Get
            Dim dr As DataRow
            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                Return (Fecha >= CDate(dr("constrdat_0")) AndAlso Fecha <= CDate(dr("conenddat_0")))

            Else
                Return False

            End If
        End Get
    End Property
    Public Property Unificacion() As Integer
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Dim i = CInt(dr("xunifi_0"))
            If i > 0 Then i = 12 \ i
            Return i
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow = dt1.Rows(0)
            dr.BeginEdit()
            dr("xunifi_0") = IIf(value >= 12, 0, value)
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property Cliente() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("conbpc_0").ToString
        End Get
    End Property
    Public Property Sucursal() As String
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return dr("bpaadd_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("bpaadd_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property Sucursales() As DataTable
        Get
            Return dt3
        End Get
    End Property
    Public ReadOnly Property FechaHasta() As Date
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return CDate(dr("conenddat_0"))
        End Get
    End Property
    Public ReadOnly Property Numero() As String
        Get
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                Return dr("connum_0").ToString

            Else
                Return ""

            End If

        End Get
    End Property
    Public Property Nombre() As String
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return dr("connam_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("connam_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property OrdenCompra() As String
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return dr("yref_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("yref_0") = IIf(value.Trim = "", " ", value)
            dr.EndEdit()
        End Set
    End Property
    Public Property Categoria() As String
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return dr("CONTYPCLA_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("CONTYPCLA_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Garantia() As Integer
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CInt(dr("gua_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("gua_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Categoria2() As Integer
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CInt(dr("concat_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("concat_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property FechaSuscripcion() As Date
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CDate(dr("siudat_0"))
        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("siudat_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property FechaInicio() As Date
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CDate(dr("constrdat_0"))
        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("constrdat_0") = value
            dr("oconstrdat_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property FechaFin() As Date
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CDate(dr("CONENDDAT_0"))
        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("CONENDDAT_0") = value
            dr("OCONENDDAT_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Duracion() As Long
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CInt(dr("rewfry_0"))
        End Get
        Set(ByVal value As Long)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("rewfry_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property DuracionBase() As Integer
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CInt(dr("REWFRYBAS_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("REWFRYBAS_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property FrecuenciaFacturacion() As Integer
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CInt(dr("INVFRY_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("INVFRY_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property FrecuenciaFacturacionBase() As Integer
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CInt(dr("INVFRYBAS_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("INVFRYBAS_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property ModoFacturacion() As Integer
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CInt(dr("INVMET_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("INVMET_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Coeficiente() As Integer
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CInt(dr("INVFRYCOE_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("INVFRYCOE_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property PreAvisoFacturacion() As Integer
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CInt(dr("INVLTI_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("INVLTI_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property PreAvisoFacturacionBase() As Integer
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CInt(dr("INVLTIBAS_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("INVLTIBAS_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property CondicionesPago() As String
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return dr("PTE_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("PTE_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property TipoCambio() As Integer
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CInt(dr("CONCHGTYP_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("CONCHGTYP_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property TipoPrecio() As Integer
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CInt(dr("CONPRITYP_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("CONPRITYP_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property CodigoArticulo() As String
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return dr("ITMREF_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("ITMREF_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property FrecuenciaReevaluacion() As Integer
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CInt(dr("RVAFRY_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("RVAFRY_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property FrecuenciaReevaluacionBase() As Integer
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CInt(dr("RVAFRYBAS_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("RVAFRYBAS_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property MetodoRevalorizacion() As Integer
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CInt(dr("RVAMET_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("RVAMET_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property MetodoRevalorizacionBase() As Integer
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CInt(dr("RVASSP_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("RVASSP_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Importe() As Double
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CDbl(dr("CONAMT_0"))
        End Get
        Set(ByVal value As Double)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("CONAMT_0") = value
            dr("OCONAMT_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Divisa() As String
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return dr("CUR_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("CUR_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property UsuarioCreacion() As String
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return dr("creusr_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("creusr_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property FechaProximaFactua() As Date
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CDate(dr("nexinvdat_0"))
        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("nexinvdat_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property FechaEnvioProximaFactura() As Date
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CDate(dr("NEXSHIINV_0"))
        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("NEXSHIINV_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property MontoProximaFactura() As Double
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CDbl(dr("nexinvamt_0"))
        End Get
        Set(ByVal value As Double)
            Dim dr As DataRow
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("nexinvamt_0") = value
            dr.EndEdit()
        End Set
    End Property

End Class