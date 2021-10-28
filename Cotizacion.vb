Imports System.Data.OracleClient
Imports System.Windows.Forms
Imports System.IO
Imports System.Collections

Public Class Cotizacion
    Public PORCENTAJE_DESCUENTO_AUTORIZADO As Double = 0.9

    Public IMPORTE_MINIMO_PEDIDO As Double = 12000 'Usado para autorizacion automatica
    'Minimo para H con entrega georgia
    Public IMPORTE_MINIMO_PEDIDO_H1 As Double = 6000
    'Minimo para H retira cliente
    Public IMPORTE_MINIMO_PEDIDO_H2 As Double = 2000
    'Importe MAXIMO permitido para LIA
    Public IMPORTE_MAXIMO_LIA As Double = 9999999999

    Private cn As OracleConnection
    Private dah As OracleDataAdapter 'Adaptador Tabla Cabecera
    Private dad As OracleDataAdapter 'Adaptador Tabla Detalle
    Private ds As New DataSet
    Private dth As DataTable 'Tabla cabecera
    Private WithEvents dtd As DataTable 'Tabla detalle
    Private dtp As DataTable 'Tabla precios
    Private bpc As Cliente
    Private bpa As Sucursal
    Private exp As Expreso
    Private tar As Tarifa

    Private Sqh As Presupuesto = Nothing
    Private Soh As Pedido = Nothing

    Public Event PresupuestoModificado(ByVal sender As Object)

    'CONSTRUCTORES
    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String

        Me.cn = cn

        'Adaptador de tabla cabecera
        Sql = "SELECT * FROM xcotiza WHERE nro_0 = :nro_0"
        dah = New OracleDataAdapter(Sql, cn)
        dah.SelectCommand.Parameters.Add("nro_0", OracleType.Number)
        dah.InsertCommand = New OracleCommandBuilder(dah).GetInsertCommand
        dah.UpdateCommand = New OracleCommandBuilder(dah).GetUpdateCommand
        dah.DeleteCommand = New OracleCommandBuilder(dah).GetDeleteCommand

        'Adaptador de tabla detalle
        Sql = "SELECT * FROM xcotizad WHERE nro_0 = :nro_0"
        dad = New OracleDataAdapter(Sql, cn)
        dad.SelectCommand.Parameters.Add("nro_0", OracleType.Number)
        dad.InsertCommand = New OracleCommandBuilder(dad).GetInsertCommand
        dad.UpdateCommand = New OracleCommandBuilder(dad).GetUpdateCommand
        dad.DeleteCommand = New OracleCommandBuilder(dad).GetDeleteCommand

        dth = New DataTable
        dtd = New DataTable
        dah.FillSchema(dth, SchemaType.Mapped)
        dad.FillSchema(dtd, SchemaType.Mapped)

        ds.Tables.Add(dth)
        ds.Tables.Add(dtd)

        tar = New Tarifa(cn)

    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal bpc As Cliente, ByVal cpy As String, ByVal EsH As Boolean, ByVal tipo As Integer)
        Me.cn = cn
        Nuevo(tipo)
    End Sub

    'SUB-PROGRAMAS
    Public Sub Nuevo(ByVal Tipo As Integer)
        Dim dr As DataRow

        dtd.Clear()
        dth.Clear()

        dr = dth.NewRow
        dr("nro_0") = 0
        dr("typ_0") = Tipo
        dr("bpcnum_0") = " "
        dr("bpaadd_0") = " "
        dr("dat_0") = Date.Today
        dr("cpy_0") = " " 'cpy
        dr("creusr_0") = USER
        dr("estado_0") = 1
        dr("h_0") = 1
        dr("soh_0") = " "
        dr("xfcrto_0") = 1
        dr("mdl_0") = 0
        dr("obs_0") = " "
        dr("yhdesde1_0") = "0000"
        dr("yhdesde2_0") = "0000"
        dr("yhhasta1_0") = "0000"
        dr("yhhasta2_0") = "0000"
        dr("valusr_0") = " "
        dr("valdat_0") = #12/31/1599#
        dr("cusordref_0") = " "
        dr("sqh_0") = " "
        dr("pte_0") = " " '"001"
        dr("ocf_0") = " "
        dr("xduplica_0") = 1
        dr("xnum_ant_0") = 0
        dr("tranum_0") = " "
        dr("licitatyp_0") = 0
        dr("licitanum_0") = " "
        dr("xitnrecha_0") = " "
        dr("xvtaantes_0") = 0
        dr("tcambio_0") = 0 'tar.CotizacionDolar
        dr("xmunro_0") = 1
        dth.Rows.Add(dr)

        bpc = Nothing
        bpa = Nothing

        Sqh = Nothing
        Soh = Nothing

    End Sub
    Public Sub Abrir(ByVal Nro As Long)
        dth.Clear()
        dtd.Clear()
        dah.SelectCommand.Parameters("nro_0").Value = Nro
        dad.SelectCommand.Parameters("nro_0").Value = Nro

        dah.Fill(dth)
        dad.Fill(dtd)

        If bpc Is Nothing Then bpc = New Cliente(cn)
        If bpa Is Nothing Then bpa = New Sucursal(cn)

        bpc.Abrir(ClienteCodigo)
        bpa.Abrir(ClienteCodigo, SucursalCodigo)

        If ExpresoCodigo.TRIM = "" Then
            exp = Nothing

        Else
            Dim dr As DataRow = dth.Rows(0)
            If exp Is Nothing Then exp = New Expreso(cn)
            exp.Abrir(dr("tranum_0").ToString)
        End If

        If PresupuestoAdonix <> "" Then
            If Sqh Is Nothing Then Sqh = New Presupuesto(cn)
            Sqh.Abrir(PresupuestoAdonix)
        End If

        If PedidoAdonix <> "" Then
            If Soh Is Nothing Then Soh = New Pedido(cn)
            Soh.Abrir(PedidoAdonix)
        End If
    End Sub
    Public Function Borrar() As Boolean

        Try
            For Each dr As DataRow In dth.Rows
                dr.Delete()
            Next

            For Each dr As DataRow In dtd.Rows
                dr.Delete()
            Next

            dah.Update(dth)
            dad.Update(dtd)

        Catch ex As Exception
            dth.RejectChanges()
            dtd.RejectChanges()

            Return False

        End Try

        Return True

    End Function
    Public Sub Duplicar()
        Dim dr As DataRow
        Dim i As Integer

        If Me.HasChanges Then Grabar()

        ds.AcceptChanges()

        For i = 0 To dtd.Rows.Count - 1
            If i = 0 Then
                dr = dth.Rows(0)

                dr.SetAdded()

                dr.BeginEdit()
                dr("nro_0") = 0
                dr("dat_0") = Date.Today
                dr("sqh_0") = " "
                dr("soh_0") = " "
                dr("typ_0") = "0"
                dr("xduplica_0") = 1
                dr.EndEdit()

            End If

            dr = dtd.Rows(i)

            dr.SetAdded()

            dr.BeginEdit()
            dr("nro_0") = 0
            dr.EndEdit()

        Next

        RecalcularPrecios()

    End Sub
    Public Function Grabar() As Boolean
        Dim n As Long = Numero 'Numero pedido actual
        Dim i As Integer = 0
        Dim dr As DataRow

        If Errores.Count > 0 Then Return False

        'Recupero el numero del pedido actual
        dr = dth.Rows(0)
        dr.BeginEdit()
        If n = 0 Then
            n = ProximoNumero()
            dr("nro_0") = n
        End If
        dr("estado_0") = IIf(Alertas(Tipo).Count > 0, 1, 2)
        dr.EndEdit()

        For Each dr In dtd.Rows
            If dr.RowState = DataRowState.Deleted Then Continue For
            i += 1000
            dr.BeginEdit()
            dr("nro_0") = n
            dr("numlig_0") = i
            dr.EndEdit()
        Next

        Try
            dah.Update(dth)
            dad.Update(dtd)

        Catch ex As OracleException
            If ex.Code = 1 Then
                dr = dth.Rows(0)
                dr.BeginEdit()
                dr("nro_0") = 0
                dr.EndEdit()
            End If

            Return False

        End Try

        Return True

    End Function
    Public Sub AceptarCambios()
        ds.AcceptChanges()
    End Sub
    Public Sub EnlazarGrilla(ByVal dgv As DataGridView)
        dgv.DataSource = dtd
    End Sub
    Private Sub RecalcularPrecios()
        Dim dr As DataRow
        Dim Art As String
        Dim Qty As Double
        Dim p As Double

        For Each dr In dtd.Rows
            Art = dr("itmref_0").ToString
            Qty = CDbl(dr("qty_0"))

            p = Precio(Art, Qty)

            dr.BeginEdit()
            'dr("precio_0") = p
            dr("precio_1") = p
            'dr("total_0") = p * Qty

            dr.EndEdit()
        Next

    End Sub
    Public Function ConvertirEnPresupuesto() As Presupuesto
        Sqh = New Presupuesto(cn)
        sqh.Nuevo(Me)
        sqh.Grabar()

        Dim dr As DataRow = dth.Rows(0)
        dr.BeginEdit()
        dr("sqh_0") = sqh.Numero
        dr("typ_0") = 2
        dr.EndEdit()

        dah.Update(dth)

        Return sqh
    End Function
    Public Function ConvertirEnPedido() As Pedido
        If ErroresPedido.Count > 0 Then Return Nothing

        Soh = New Pedido(cn)
        Soh.Nuevo(Me)
        Soh.Grabar()

        Dim dr As DataRow = dth.Rows(0)
        dr.BeginEdit()
        dr("soh_0") = Soh.Numero
        dr("typ_0") = 1
        dr("valusr_0") = USER
        dr("valdat_0") = Date.Today
        dr.EndEdit()
        dah.Update(dth)

        Return Soh

    End Function
    Public Sub SetCliente(Optional ByVal Codigo As String = "")
        Dim dr As DataRow

        If Codigo = "" Then
            bpc = Nothing
            SetSucursal()
            'exp = Nothing
        Else
            If bpc Is Nothing Then bpc = New Cliente(cn)
            bpc.Abrir(Codigo)
            SetSucursal()
            'exp = Nothing
        End If

        If dth.Rows.Count = 1 Then
            dr = dth.Rows(0)
            dr.BeginEdit()
            'Tranfiere a la cotizacion los datos del cliente
            If bpc IsNot Nothing Then
                dr("bpcnum_0") = bpc.Codigo
                dr("xfcrto_0") = bpc.FcRto
                dr("pte_0") = bpc.CondicionDePago
                dr("xvtaantes_0") = IIf(bpc.AVentasAntesEntregar, 2, 1)
            Else
                dr("bpcnum_0") = " "
                dr("xfcrto_0") = 1
                dr("pte_0") = " "
                dr("xvtaantes_0") = 1
            End If
            dr.EndEdit()
        End If

    End Sub
    Public Sub SetSucursal(Optional ByVal Codigo As String = "")
        Dim dr As DataRow

        If Codigo = "" Then
            bpa = Nothing
            exp = Nothing
            SetExpreso()
        Else
            If bpa Is Nothing Then bpa = New Sucursal(cn)
            bpa.Abrir(Me.Cliente.Codigo, Codigo)
            SetExpreso(bpa.Expreso)
        End If

        If dth.Rows.Count = 1 Then
            dr = dth.Rows(0)
            dr.BeginEdit()
            'Tranfiere a la cotizacion los datos de la sucursal
            If bpa IsNot Nothing Then
                dr("bpaadd_0") = bpa.Sucursal
                'dr("mdl_0") = bpa.ModoEntrega
                dr("yhdesde1_0") = bpa.TurnoMananaDesde
                dr("yhhasta1_0") = bpa.TurnoMananaHasta
                dr("yhdesde2_0") = bpa.TurnoTardeDesde
                dr("yhhasta2_0") = bpa.TurnoTardeHasta
                dr("xmunro_0") = IIf(bpa.SaleDesdeMunro, 2, 1)
            Else
                dr("bpaadd_0") = " "
                'dr("mdl_0") = "0"
                dr("yhdesde1_0") = " "
                dr("yhdesde2_0") = " "
                dr("yhhasta1_0") = " "
                dr("yhhasta2_0") = " "
                dr("xmunro_0") = 0
            End If
            dr.EndEdit()
        End If
    End Sub
    Public Sub SetExpreso(Optional ByVal Codigo As String = "")
        Codigo = Codigo.Trim

        If Codigo = "" Then
            exp = Nothing
        Else
            If exp Is Nothing Then exp = New Expreso(cn)
            exp.Abrir(Codigo)
        End If

        Dim dr As DataRow = dth.Rows(0)
        dr.BeginEdit()
        If exp Is Nothing Then
            dr("tranum_0") = " "
        Else
            dr("tranum_0") = exp.Codigo
        End If
        dr.EndEdit()

    End Sub
    Public Function TieneArticulosTipo(ByVal Tipo As Integer) As Boolean
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable

        Sql = "SELECT xcd.* "
        Sql &= "FROM xcotizad xcd INNER JOIN xprecios xpe ON (xcd.itmref_0 = xpe.itmref_0) "
        Sql &= "WHERE ped_0 = :tipo AND nro_0 = :nro and xcd.itmref_0 <> 653001"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("nro", OracleType.Number).Value = Numero
        da.SelectCommand.Parameters.Add("tipo", OracleType.Number).Value = Tipo

        da.Fill(dt)
        da.Dispose()

        Return dt.Rows.Count > 0

    End Function
    
    Public Function TieneLineaDeProducto(ByVal flia As String, ByVal flia2 As String, ByVal num As Boolean) As Boolean
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable

        Sql = "SELECT xcd.* "
        Sql &= "FROM xcotizad xcd INNER JOIN "
        Sql &= "     xprecios xpe ON (xcd.itmref_0 = xpe.itmref_0) inner join "
        Sql &= "     itmmaster itmm on (xcd.itmref_0 = itmm.itmref_0) and (cfglin_0 >= :flia and cfglin_0 <= :flia2) "
        Sql &= "WHERE nro_0 = :nro and xcd.itmref_0 <> '653001' "

        If num Then Sql &= "and numlig_0  = 1000 "

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("nro", OracleType.Number).Value = Numero
        da.SelectCommand.Parameters.Add("flia", OracleType.VarChar).Value = flia
        da.SelectCommand.Parameters.Add("flia2", OracleType.VarChar).Value = flia2

        da.Fill(dt)
        da.Dispose()

        Return dt.Rows.Count > 0

    End Function
    Public Sub AgregarLinea(ByVal Articulo As String, ByVal Cantidad As Double, Optional ByVal Reemplazar As Boolean = False)
        Dim dr As DataRow
        Dim p As Double = Precio(Articulo, Cantidad)

        If Reemplazar Then
            For Each dr In dtd.Rows
                If dr("itmref_0").ToString = Articulo Then
                    dr.BeginEdit()
                    'dr("nro_0") = Numero
                    'dr("numlig_0") = dtd.Rows.Count + 1
                    'dr("itmref_0") = Articulo
                    dr("qty_0") = Cantidad
                    'dr("precio_0") = p
                    'dr("precio_1") = p
                    dr("total_0") = p * Cantidad
                    dr.EndEdit()
                    Exit Sub
                End If
            Next
        Else
            For Each dr In dtd.Rows
                If dr("itmref_0").ToString = Articulo Then
                    'Agrego a la cantidad el valor anterior
                    Cantidad += CDbl(dr("qty_0"))

                    dr.BeginEdit()
                    'dr("nro_0") = Numero
                    'dr("numlig_0") = dtd.Rows.Count + 1
                    'dr("itmref_0") = Articulo
                    dr("qty_0") = Cantidad
                    'dr("precio_0") = p
                    'dr("precio_1") = p
                    dr("total_0") = p * Cantidad
                    dr.EndEdit()
                    Exit Sub
                End If
            Next
        End If

        dr = dtd.NewRow
        dr("nro_0") = Numero
        dr("numlig_0") = dtd.Rows.Count + 1
        dr("itmref_0") = Articulo
        dr("qty_0") = Cantidad
        dr("precio_0") = p
        dr("precio_1") = p
        dr("total_0") = p * Cantidad
        dtd.Rows.Add(dr)

    End Sub

    'FUNCIONES
    'Modo = 1 Pedido
    'Modo = 2 Presupuesto
    Public Function Alertas(Optional ByVal Modo As Integer = 2) As ArrayList
        Dim txt As String = ""
        Dim _Alertas As New ArrayList

        _Alertas.Clear()

        '2021-08-02 No se valida para presupuestos
        If Modo = 2 Then Return _Alertas

        'Chequeo si los precios no superan el PORCENTAJE AUTORIZADO DE DESCUENTO
        For Each dr As DataRow In dtd.Rows
            Dim p1, p2 As Double

            If dr.RowState = DataRowState.Deleted Then Continue For

            p1 = CDbl(dr("precio_0")) 'Precio vendedor
            p2 = CDbl(dr("precio_1")) 'Precio sugerido
            p2 = p2 * PORCENTAJE_DESCUENTO_AUTORIZADO

            If Not (p1 >= p2) Then
                txt = dr("itmref_0").ToString
                txt &= " supera descuento autorizado "
                txt &= "(" & (1 - PORCENTAJE_DESCUENTO_AUTORIZADO) * 100 & "%)"

                _Alertas.Add(txt)
            End If

            'Chequeo que para articulos 359026 y 359027 la cantidad sea multiplo de 18
            If dr("itmref_0").ToString = "359026" Or dr("itmref_0").ToString = "359027" Then
                If Not (CInt(dr("qty_0")) Mod 18 = 0) Then
                    txt = dr("itmref_0").ToString
                    txt &= " - cantidad debe ser múltiplo de 18"
                    _Alertas.Add(txt)
                End If
            End If
        Next

        FaltanArticulosRelacionados(_Alertas)

        If ModoEntrega = "1" AndAlso H Then
            If PrecioTotalAI < IMPORTE_MINIMO_PEDIDO_H1 And bpc.EsAbonado = False AndAlso bpc.TipoAbcStr.Contains("A") = False Then
                txt = "-> Importe mínimo para H y Entrega Georgia es " & IMPORTE_MINIMO_PEDIDO_H1.ToString("N2")
                _Alertas.Add(txt)
            End If

        ElseIf ModoEntrega <> "1" AndAlso H Then
            If PrecioTotalAI < IMPORTE_MINIMO_PEDIDO_H2 AndAlso bpc.EsAbonado = False AndAlso bpc.TipoAbcStr.Contains("A") = False Then
                txt = "-> Importe mínimo para H y Retira Cliente es " & IMPORTE_MINIMO_PEDIDO_H2.ToString("N2")
                _Alertas.Add(txt)
            End If

        ElseIf ModoEntrega = "1" Then
            Dim itn As New Intervencion(cn)
            If IntervencionRechazo.ToString <> " " AndAlso itn.Abrir(IntervencionRechazo.ToString) Then
                If itn.Sector = "CTD" Or itn.Sector = "ADM" Or itn.Sector = "SRV" Then
                    'No validar nada
                Else
                    If PrecioTotalAI < IMPORTE_MINIMO_PEDIDO AndAlso _
                       bpc.EsAbonado = False AndAlso _
                       bpc.TipoAbcStr.Contains("A") = False AndAlso _
                       ExisteFleteYAcarreo() = False Then

                        txt = "-> Importe mínimo para Entrega Georgia es " & IMPORTE_MINIMO_PEDIDO.ToString("N2")
                        _Alertas.Add(txt)
                    End If
                End If
            Else
                If PrecioTotalAI < IMPORTE_MINIMO_PEDIDO AndAlso _
                   bpc.EsAbonado = False AndAlso _
                   bpc.TipoAbcStr.Contains("A") = False AndAlso _
                   ExisteFleteYAcarreo() = False Then

                    txt = "-> Importe mínimo para Entrega Georgia es " & IMPORTE_MINIMO_PEDIDO.ToString("N2")
                    _Alertas.Add(txt)
                End If
            End If
        End If

        'Verificacion de prioridad de condicion de pago
        If Not Cliente.EsProspecto Then
            Dim cp As New CondicionPago(cn)

            cp.EstablecerCodigo(CondicionPago)

            If Cliente.CondicionPago.Prioridad < cp.Prioridad Then
                txt = "-> Condición de pago modificada es desfavorable"
                _Alertas.Add(txt)
            End If

            If Modo = 1 AndAlso bpa.Portero.Trim = "" Then
                txt = "-> Nombre de contacto obligatorio"
                _Alertas.Add(txt)
            End If
            If Modo = 1 AndAlso Not Utiles.ValidarTelefono(bpa.Telefono_Portero) Then
                txt = "-> Teléfono de contacto inválido"
                _Alertas.Add(txt)
            End If
        End If

        'Pedido / Entrega Georgia / Direccion Entrega
        If Modo = 1 AndAlso ModoEntrega = "1" AndAlso bpa.EsDireccionEntrega Then
            Dim e As String = ""

            If Not Utiles.ValidarFranjasHorarias(HoraMananaDesde, HoraMananaHasta, HoraTardeDesde, HoraTardeHasta, e) Then
                _Alertas.Add("-> " & e)
            End If
        End If

        'Mantenimientos en el interior del pais
        If bpa.Provincia <> "CFE" And bpa.Provincia <> "BUE" Then
            If Me.TieneArticulo("551015") Or Me.TieneArticulo("551016") Then
                _Alertas.Add("-> Mantenimientos en el interior debe ser aprobado por supervisor")
            End If
        End If

        'No se puede crear pedido, si el presupuesto está vencido
        If PresupuestoAdonix.Trim <> "" Then
            Dim sqh As New Presupuesto(cn)

            If sqh.Abrir(PresupuestoAdonix) Then
                If sqh.Vencimiento.AddDays(15) <= Date.Today Then
                    _Alertas.Add("-> El presupuesto está vencido")
                Else
                    'Elimino Alertas para crear pedido
                    If sqh.Vencimiento < Date.Today Then _Alertas.Clear()
                End If
            End If
        End If

        If Modo = 1 AndAlso Cliente.OC_obligatoria AndAlso Me.OC = " " Then
            _Alertas.Add("-> OC obligatoria para éste cliente.")
        End If

        If Modo = 1 AndAlso PresupuestoAdonix.Trim <> "" Then

            If Me.Presupuesto.TipoCambio > 0 AndAlso Me.Presupuesto.TipoCambio * 1.05 < tar.CotizacionDolar Then
                txt = "La diferencia de cambio es superior al 5%" & vbCrLf & vbCrLf
                txt &= Me.Presupuesto.TipoCambio.ToString("N2") & " <==> " & TipoCambio.ToString("N2")
                _Alertas.Add(txt)
            End If

        End If

        Return _Alertas

    End Function
    Private Function TieneArticulosTipoCotizacion() As Boolean
        Dim valor As Boolean = False
        For Each dr As DataRow In dtd.Rows
            If dr.RowState = DataRowState.Deleted Then Continue For
            valor = TieneArticulosTipoCotizacion(CInt(dr("itmref_0")))
        Next
        Return valor
    End Function
    Public Function TieneArticulosTipoCotizacion(ByVal item As Integer) As Boolean
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable

        Sql = "SELECT * FROM xprecios "
        Sql &= " WHERE ped_0 = 2 AND itmref_0 = :nro and itmref_0 <> 653001"


        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("nro", OracleType.Number).Value = item
        '  da.SelectCommand.Parameters.Add("tipo", OracleType.Number).Value = 2

        da.Fill(dt)
        da.Dispose()

        Return dt.Rows.Count > 0

    End Function
    Public Function Errores() As ArrayList
        Dim itm As New Articulo(cn)
        Dim _Errores As New ArrayList

        _Errores.Clear()

        If bpc Is Nothing Then
            _Errores.Add("Cliente no definido")
            Return _Errores
        End If
        If bpa Is Nothing Then
            _Errores.Add("Sucursal no definida")
            Return _Errores
        End If

        If Me.Tipo = 0 Then _Errores.Add("Presupuesto/Pedido no definido")
        If Me.SociedadCodigo.Trim = "" Then _Errores.Add("Falta seleccionar sociedad")
        If Me.ModoEntrega = "0" Then _Errores.Add("Falta Modo de Entrega")
        If Me.OC.Trim <> "" AndAlso Me.OCF.Trim = "" Then _Errores.Add("Falta subir archivo con OC")
        If Me.CondicionPago.Trim = "" Then _Errores.Add("Falta seleccionar Condición de Pago")

        'Usadas para no repetir los errores
        Dim v1 As Boolean = True
        Dim v2 As Boolean = True
        Dim v3 As Boolean = True
        Dim i As Integer = 0 'Usada para contar las lineas

        'VALIDACION 1 - CLIENTE DISTRIBUIDOR Y ARTICULOS 10 Y 11
        'VALIDACION 2 - CANTIDADES EN CERO
        For Each dr As DataRow In dtd.Rows
            'Abro el articulo
            If dr.RowState = DataRowState.Deleted Then Continue For
            If Not itm.Abrir(dr("itmref_0").ToString) Then Continue For

            i += 1

            If itm.Familia(1) = "101" Or itm.Familia(1) = "102" Or itm.Familia(1) = "104" Then
                'Matafuego de 1Kg se debe cargar con codigo 10xxxx
                If Not itm.Codigo.StartsWith("10") Then
                    If v1 Then
                        _Errores.Add("Extintores de 1KG se deben cargar con código 10")
                        v1 = False
                    End If
                End If
            Else
                'A distribuidores se debe cargar con 10xxxx
                If bpc.Familia2 = "20" AndAlso itm.Codigo.StartsWith("11") Then
                    'Como el cliente es DISTRIBUIDOR - Articulos NO pueden ser 11XXXX
                    If v2 Then
                        _Errores.Add("CLIENTE DISTRIBUIDOR: No se pueden cargar artículos 11")
                        v2 = False
                    End If

                ElseIf bpc.Familia2 <> "20" AndAlso itm.Codigo.StartsWith("10") Then
                    'Cliente NO DISTRIBUIDOR - Articulos NO pueden ser 10XXXX
                    If bpc.Familia2 <> "30" Then 'No aplica para supermercados
                        If v3 Then
                            _Errores.Add("CLIENTE NO DISTRIBUIDOR: No se pueden cargar artículos 10")
                            v3 = False
                        End If
                    End If
                End If
            End If

            If CDbl(dr("qty_0")) = 0 Then
                _Errores.Add("No se permiten cantidades en cero")
                Exit For
            End If

        Next

        If i = 0 Then _Errores.Add("No hay líneas cargadas")

        If bpc.EmpresaFacturacionObligatoria <> "" Then

            If bpc.EmpresaFacturacionObligatoria <> Me.SociedadCodigo Then
                _Errores.Add("Este cliente se factura solo en: " & bpc.EmpresaFacturacionObligatoria)
            End If

        End If

        Select Case Me.SociedadCodigo
            Case "MON"
                If Not Me.H Then _Errores.Add("La sociedad MON solo se puede usar en H")

            Case "SCH"
                If bpc.RegimenImpuesto = "RI" OrElse bpc.RegimenImpuesto = "RIE" Then
                    _Errores.Add("La sociedad SCH no puede usarse con clientes RI")
                End If

            Case "GRU", "LIA"
                _Errores.Add("No se puede usar la sociedad " & Me.SociedadCodigo)

        End Select

        itm.Dispose()
        itm = Nothing

        Return _Errores

    End Function
    Public Function ErroresPedido() As ArrayList
        Dim txt As String = ""
        Dim _Errores As New ArrayList

        _Errores.Clear()

        If bpc.EsProspecto Then
            _Errores.Add("No se pude crear un pedido a un prospecto")
            Return _Errores
        End If

        If SociedadCodigo = "DNY" Then
            With bpc
                If .RegimenImpuesto = "RI" AndAlso .IIBB = "" And .CondicionIb <> "4" Then
                    _Errores.Add("IIBB obligatorio")
                End If
            End With
        End If
        If Not TieneArticulosTipo(2) Then
            _Errores.Add("Esta cotizacion no tiene articulos nuevos")
        End If
        If bpa.SucursalEntregaActiva = False Then
            _Errores.Add("La sucursal no es para entregas o está desactivada")
        End If
        'Expreso obligatorio para los clientes distinto a BUE y CFE
        If ExpresoCodigo.Trim = "" AndAlso _
           ModoEntrega = "1" AndAlso _
           bpa.Provincia <> "BUE" AndAlso _
           bpa.Provincia <> "CFE" AndAlso _
           TieneArticulo("551016") = False Then

            _Errores.Add("Expreso obligatorio")

        End If

        If FcRto AndAlso CondicionPago > "001" And CondicionPago <= "023" Then
            _Errores.Add("Factura con Rto no permite condición contra entrega")
        End If
        If CondicionPago = "062" Then
            _Errores.Add("Condición de pago no permitida para pedidos")
        End If

        If SociedadCodigo = "SCH" And PrecioTotalII > IMPORTE_MAXIMO_LIA Then
            txt = "No se pude facturar más de {0} en sociedad {1}"
            txt = txt.Replace("{0}", IMPORTE_MAXIMO_LIA.ToString("N2"))
            txt = txt.Replace("{1}", SociedadCodigo)
            _Errores.Add(txt)
        End If



        Return _Errores

    End Function
    Public Function CambioPrecio() As Boolean
        Dim dr As DataRow

        For Each dr In dtd.Rows
            If Precio(dr("itmref_0").ToString, CDbl(dr("qty_0"))) > CDbl((dr("precio_0"))) Then
                Return True
            End If
        Next
        Return False
    End Function
    Private Function ProximoNumero() As Long
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim Nro As Long = 0

        da = New OracleDataAdapter("SELECT MAX(nro_0) FROM xcotiza", cn)
        da.Fill(dt)
        da.Dispose()

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            If Not IsDBNull(dr(0)) Then Nro = CLng(dr(0))
        End If
        Nro += 1
        Return Nro

    End Function
    Public Function TieneArtComenzandoCon(ByVal Str As String) As Boolean
        Dim dr As DataRow
        Dim flg As Boolean = False

        For Each dr In dtd.Rows
            'Salto si el registro fue eliminado
            If dr.RowState = DataRowState.Deleted Then Continue For

            If dr("itmref_0").ToString.StartsWith(Str) Then
                flg = True
                Exit For
            End If

        Next

        Return flg

    End Function
    Public Function TieneArticulo(ByVal Str As String) As Boolean
        Dim dr As DataRow
        Dim flg As Boolean = False

        For Each dr In dtd.Rows
            'Salto si el registro fue eliminado
            If dr.RowState = DataRowState.Deleted Then Continue For

            If dr("itmref_0").ToString = Str Then
                flg = True
                Exit For
            End If

        Next

        Return flg

    End Function
    '*************************************************************************
    'Funcion que devuelve los articulos que están relacionados con otro
    '*************************************************************************
    Public Function ArticulosRelacionados(ByVal Codigo As String) As DataTable
        Dim Sql As String = "SELECT * FROM xcotizax WHERE ori_0 = :ori"
        Dim da As OracleDataAdapter
        Dim dt As New DataTable

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("ori", OracleType.VarChar).Value = Codigo
        da.Fill(dt)

        'Agrego todos los articulos relacionados al array
        'For Each dr As DataRow In dt.Rows
        '    arr.Add(New Articulo(cn, dr("des_0").ToString))
        'Next

        da.Dispose()
        dt.Dispose()

        Return dt

    End Function
    Public Function RelacionTarjetasExtintores() As Boolean
        'Devuelve TRUE si la cantidad de tarjetas es igual a la cantidad de extintores
        Dim flg As Boolean = True
        Dim cant_tarj As Integer = 0
        Dim cant_ext As Integer = 0
        Dim itm As New Articulo(cn)

        'Recorro las lineas y sumo las tarjetas y extintores
        For Each dr As DataRow In dtd.Rows
            If dr.RowState = DataRowState.Deleted Then Continue For

            'Sumo cantidad de tarjetas
            If dr("itmref_0").ToString = "705020" Then
                cant_tarj += CInt(dr("qty_0"))
                Continue For
            End If

            'Sumo cantidad de extintores
            If itm.Abrir(dr("itmref_0").ToString) Then
                If itm.Categoria = "10" AndAlso itm.Familia(1) <> "101" AndAlso itm.Familia(1) <> "102" Then
                    cant_ext += CInt(dr("qty_0"))
                End If
            End If
        Next

        Return cant_ext = cant_tarj

    End Function
    Public Function TieneArticuloBUE() As Boolean

        Dim flg As Boolean = False
        Dim itm As New Articulo(cn)

        'Recorro las lineas y sumo las tarjetas y extintores
        For Each dr As DataRow In dtd.Rows
            If dr.RowState = DataRowState.Deleted Then Continue For

            'Sumo cantidad de tarjetas
            If dr("itmref_0").ToString = "705020" Then
                flg = True
                Continue For
            End If
        Next
        Return flg

    End Function
    Public Function TieneArticulosCFE() As Boolean
        'Devuelve TRUE si no encuentra Si el pedido está cargado a un domicilio CFE, tiene al menos un articulo con linea de producto entre 102 y 111 incluidos,  y no tiene una linea cargada por 705020
        Dim flg As Boolean = False
       
        Dim itm As New Articulo(cn)

        'Recorro las lineas y sumo las tarjetas y extintores
        For Each dr As DataRow In dtd.Rows
            If dr.RowState = DataRowState.Deleted Then Continue For
            If itm.Abrir(dr("itmref_0").ToString) Then
                If (CInt(itm.Familia(1)) >= 102 And CInt(itm.Familia(1)) <= 111) Then
                    For Each dr2 As DataRow In dtd.Rows
                        If Not dr2("itmref_0").ToString = "705020" Then
                            flg = True
                            Continue For
                        Else
                            flg = False
                            Exit For
                        End If
                    Next
                End If
            End If
        Next

        Return flg

    End Function
    Public Function Extintores3pulgadas() As Boolean
        'Da falso si existen cantidades de 3" mayores a 10 y que no son multiplo de 10
        Dim dr As DataRow
        Dim itm As New Articulo(cn)
        Dim flg As Boolean = True
        Dim Pax As Integer = 6

        For Each dr In dtd.Rows
            If dr.RowState = DataRowState.Deleted Then Continue For
            If Not itm.Abrir(dr("itmref_0").ToString) Then Continue For

            If itm.Familia(1) = "101" Then
                If CDbl(dr("qty_0")) > Pax AndAlso CDbl(dr("qty_0")) Mod Pax > 0 Then
                    flg = False
                    Exit For
                End If
            End If

        Next

        Return flg

    End Function
    Public Function Extintores4pulgadas() As Boolean
        'Da falso si existen cantidades de 3" mayores a 10 y que no son multiplo de 10
        Dim dr As DataRow
        Dim itm As New Articulo(cn)
        Dim flg As Boolean = True
        Dim Pax As Integer = 12

        For Each dr In dtd.Rows
            If dr.RowState = DataRowState.Deleted Then Continue For
            If Not itm.Abrir(dr("itmref_0").ToString) Then Continue For

            If itm.Familia(1) = "102" Then
                If CDbl(dr("qty_0")) > Pax AndAlso CDbl(dr("qty_0")) Mod Pax > 0 Then
                    flg = False
                    Exit For
                End If
            End If

        Next

        Return flg

    End Function
    Public Function Precio(ByVal articulo As String, ByVal cantidad As Double) As Double
        Dim dr As DataRow = Nothing
        Dim p As Double = 0
        Dim dt As New DataTable

        p = tar.ObtenerPrecio(Cliente, articulo, cantidad)

        Return p

    End Function
    Public Function ExisteArticulo(ByVal Articulo As String) As Boolean
        Return tar.ExisteArticulo(Articulo)
    End Function
    
    'PROPIEDADES
    Public ReadOnly Property Numero() As Long
        Get
            Dim dr As DataRow
            Dim n As Long = 0

            If dth.Rows.Count > 0 Then
                dr = dth.Rows(0)
                n = CLng(dr("nro_0"))
            End If

            Return n
        End Get
    End Property
    Public ReadOnly Property Cliente() As Cliente
        Get
            Return bpc
        End Get
    End Property
    Public ReadOnly Property ClienteCodigo() As String
        Get
            Dim dr As DataRow
            Dim s As String = ""

            If dth.Rows.Count > 0 Then
                dr = dth.Rows(0)
                s = dr("bpcnum_0").ToString
            End If

            Return s
        End Get
    End Property
    Public Property EsDuplicado() As Boolean
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return CBool(IIf(CInt(dr("xduplica_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow = dth.Rows(0)
            dr.BeginEdit()
            dr("xduplica_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property CotizacionOrigen() As Integer
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return CInt(dr("xnum_ant_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow = dth.Rows(0)
            dr.BeginEdit()
            dr("xnum_ant_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property Sucursal() As Sucursal
        Get
            Return bpa
        End Get
    End Property
    Public ReadOnly Property SucursalCodigo() As String
        Get
            Dim dr As DataRow
            Dim s As String = ""

            If dth.Rows.Count > 0 Then
                dr = dth.Rows(0)
                s = dr("bpaadd_0").ToString
            End If

            Return s
        End Get
    End Property
    Public ReadOnly Property Fecha() As Date
        Get
            Dim dr As DataRow
            Dim d As Date

            If dth.Rows.Count > 0 Then
                dr = dth.Rows(0)
                d = CDate(dr("dat_0"))
            End If

            Return d
        End Get
    End Property
    Public Property SociedadCodigo() As String
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return dr("cpy_0").ToString
        End Get
        Set(ByVal value As String)
            If dth.Rows.Count = 0 Then Exit Property
            If value = SociedadCodigo Then Exit Property
            Dim dr As DataRow = dth.Rows(0)

            dr.BeginEdit()
            dr("cpy_0") = value
            dr.EndEdit()

            'Fuerzo quitar H si no es MON
            H = value = "MON"

            RecalcularPrecios()

        End Set
    End Property
    Public ReadOnly Property Sociedad() As Sociedad
        Get
            Dim dr As DataRow = dth.Rows(0)
            Dim cpy As New Sociedad(cn)

            cpy.abrir(dr("cpy_0").ToString)

            Return cpy

        End Get
    End Property
    Public Property CondicionPago() As String
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return dr("pte_0").ToString
        End Get
        Set(ByVal value As String)
            If (value = CondicionPago) Then Exit Property
            If dth.Rows.Count = 0 Then Exit Property
            Dim dr As DataRow = dth.Rows(0)

            dr.BeginEdit()
            dr("pte_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property H() As Boolean
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return CBool(IIf(CInt(dr("h_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            If dth.Rows.Count = 0 Then Exit Property
            If value = H Then Exit Property
            Dim dr As DataRow = dth.Rows(0)

            dr.BeginEdit()
            dr("h_0") = IIf(value, 2, 1)
            dr.EndEdit()

            'Fuerzo MON si es H
            If value Then SociedadCodigo = "MON"

            RecalcularPrecios()

            'RaiseEvent CotizacionModificada(Me)
        End Set
    End Property
    Public ReadOnly Property UsuarioCreacion() As Usuario
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return New Usuario(cn, dr("creusr_0").ToString)
        End Get
    End Property
    Public ReadOnly Property Estado() As Integer
        Get
            Dim dr As DataRow
            Dim n As Integer = 1

            If dth.Rows.Count > 0 Then
                dr = dth.Rows(0)
                n = CInt(dr("estado_0"))
            End If

            Return n
        End Get
    End Property
    Public ReadOnly Property PedidoAdonix() As String
        Get
            Dim dr As DataRow
            Dim n As String = ""

            If dth.Rows.Count > 0 Then
                dr = dth.Rows(0)
                n = dr("soh_0").ToString.Trim
            End If

            Return n
        End Get
    End Property
    Public ReadOnly Property PresupuestoAdonix() As String
        Get
            Dim dr As DataRow
            Dim n As String = ""

            If dth.Rows.Count > 0 Then
                dr = dth.Rows(0)
                n = dr("sqh_0").ToString.Trim
            End If

            Return n
        End Get
    End Property
    Public Property Tipo() As Integer
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return CInt(dr("typ_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow = dth.Rows(0)
            dr.BeginEdit()
            dr("typ_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property FcRto() As Boolean
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return CBool(IIf(CInt(dr("xfcrto_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            If dth.Rows.Count = 0 Then Exit Property
            If value = FcRto Then Exit Property
            Dim dr As DataRow = dth.Rows(0)

            dr.BeginEdit()
            dr("xfcrto_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property ModoEntrega() As String
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return dr("mdl_0").ToString
        End Get
        Set(ByVal value As String)
            If dth.Rows.Count = 0 Then Exit Property
            If value = ModoEntrega Then Exit Property
            Dim dr As DataRow = dth.Rows(0)

            dr.BeginEdit()
            dr("mdl_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Obs() As String
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return dr("obs_0").ToString
        End Get
        Set(ByVal value As String)
            If value.Trim = "" Then value = " "
            If dth.Rows.Count = 0 Then Exit Property

            Dim dr As DataRow = dth.Rows(0)

            dr.BeginEdit()
            dr("obs_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property HoraMananaDesde() As String
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return dr("yhdesde1_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dth.Rows(0)
            dr.BeginEdit()
            dr("yhdesde1_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property HoraMananaHasta() As String
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return dr("yhhasta1_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dth.Rows(0)
            dr.BeginEdit()
            dr("yhhasta1_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property HoraTardeDesde() As String
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return dr("yhdesde2_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dth.Rows(0)
            dr.BeginEdit()
            dr("yhdesde2_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property HoraTardeHasta() As String
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return dr("yhhasta2_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dth.Rows(0)
            dr.BeginEdit()
            dr("yhhasta2_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property OC() As String
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return dr("cusordref_0").ToString
        End Get
        Set(ByVal value As String)
            If value.Trim = "" Then value = " "

            If dth.Rows.Count = 0 Then Exit Property
            Dim dr As DataRow = dth.Rows(0)

            If value = "" Then value = " "

            dr.BeginEdit()
            dr("cusordref_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property OCF() As String
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return dr("ocf_0").ToString
        End Get
        Set(ByVal value As String)
            If value.Trim = "" Then value = " "

            If dth.Rows.Count = 0 Then Exit Property
            Dim dr As DataRow = dth.Rows(0)

            If value = "" Then value = " "

            dr.BeginEdit()
            dr("ocf_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property PrecioTotalAI() As Double
        Get
            Dim dr As DataRow
            Dim p As Double = 0

            For Each dr In dtd.Rows
                If dr.RowState = DataRowState.Deleted Then Continue For
                p += CDbl(dr("total_0"))
            Next

            Return p

        End Get
    End Property
    Public ReadOnly Property PrecioTotalII() As Double
        Get
            Dim p As Double = PrecioTotalAI

            If SociedadCodigo <> "MON" Then p *= 1.21

            Return p

        End Get
    End Property
    Public ReadOnly Property HasChanges() As Boolean
        Get
            'Devuelve si el pedido tiene modificaciones no guardadas
            Return ds.HasChanges
        End Get
    End Property
    Public ReadOnly Property Lineas() As DataTable
        Get
            Return dtd
        End Get
    End Property
    Public ReadOnly Property Expreso() As Expreso
        Get
            Return exp
        End Get
    End Property
    Public ReadOnly Property ExpresoCodigo() As String
        Get
            Dim dr As DataRow
            Dim s As String = " "

            If dth.Rows.Count > 0 Then
                dr = dth.Rows(0)
                s = dr("tranum_0").ToString
            End If

            Return s
        End Get
    End Property
    Public Property TipoLicitacion() As Integer
        Get
            Dim i As Integer = 1

            If dth.Rows.Count > 0 Then
                Dim dr As DataRow = dth.Rows(0)
                i = CInt(dr("licitatyp_0"))
            End If
            If i = 0 Then i = 1
            Return i
        End Get
        Set(ByVal value As Integer)

            If dth.Rows.Count = 0 Then Exit Property
            Dim dr As DataRow = dth.Rows(0)

            dr.BeginEdit()
            dr("licitatyp_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property NumeroLicitacion() As String
        Get
            Dim txt As String = ""

            If dth.Rows.Count > 0 Then
                Dim dr As DataRow = dth.Rows(0)
                txt = dr("licitanum_0").ToString
            End If

            Return txt.Trim
        End Get
        Set(ByVal value As String)
            If dth.Rows.Count = 0 Then Exit Property
            Dim dr As DataRow = dth.Rows(0)

            dr.BeginEdit()
            dr("licitanum_0") = IIf(value.Trim = "", " ", value)
            dr.EndEdit()
        End Set
    End Property
    Public Property IntervencionRechazo() As String
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return dr("xitnrecha_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dth.Rows(0)
            dr.BeginEdit()
            dr("xitnrecha_0") = If(value.Trim = "", " ", value)
            dr.EndEdit()
        End Set
    End Property
    Public Property AVentasAntesEntregar() As Boolean
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return CBool(IIf(CInt(dr("xvtaantes_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            If dth.Rows.Count = 0 Then Exit Property
            'If value = FcRto Then Exit Property
            Dim dr As DataRow = dth.Rows(0)

            dr.BeginEdit()
            dr("xvtaantes_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property TipoCambio() As Double
        Get
            Dim dr As DataRow
            Dim p As Double = 0

            If dth.Rows.Count > 0 Then
                dr = dth.Rows(0)
                p = CDbl(dr("tcambio_0"))
            End If

            Return p
        End Get
        Set(ByVal value As Double)
            Dim dr As DataRow

            If dth.Rows.Count > 0 Then
                dr = dth.Rows(0)
                dr.BeginEdit()
                dr("tcambio_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property SaleDesdeMunro() As Boolean
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return CInt(dr("xmunro_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow = dth.Rows(0)
            dr.BeginEdit()
            dr("xmunro_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Shared Sub Exportar(ByVal cn As OracleConnection, ByVal Archivo As String)
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim sw As StreamWriter
        Dim linea As String

        Sql = "SELECT * FROM xprecios"
        da = New OracleDataAdapter(Sql, cn)

        Try
            sw = New StreamWriter(Archivo, False, System.Text.Encoding.Default)

            da.Fill(dt)

            For Each dr In dt.Rows
                linea = dr("itmref_0").ToString.Replace(".", ",") & vbTab
                linea &= dr("qty_0").ToString.Replace(".", ",") & vbTab
                linea &= dr("precio_0").ToString.Replace(".", ",") & vbTab
                linea &= dr("precio_1").ToString.Replace(".", ",") & vbTab
                linea &= dr("precio_2").ToString.Replace(".", ",") & vbTab
                linea &= dr("precio_3").ToString.Replace(".", ",") & vbTab
                linea &= dr("precio_4").ToString.Replace(".", ",") & vbTab
                linea &= dr("precio_5").ToString.Replace(".", ",") & vbTab
                linea &= dr("precio_6").ToString.Replace(".", ",") & vbTab
                linea &= dr("precio_7").ToString.Replace(".", ",") & vbTab
                linea &= dr("precio_8").ToString.Replace(".", ",") & vbTab
                linea &= dr("precio_9").ToString.Replace(".", ",") & vbTab
                linea &= dr("precio_10").ToString.Replace(".", ",") & vbTab
                linea &= dr("precio_11").ToString.Replace(".", ",") & vbTab
                linea &= dr("precio_12").ToString.Replace(".", ",") & vbTab
                linea &= dr("precio_13").ToString.Replace(".", ",") & vbTab
                linea &= dr("precio_14").ToString.Replace(".", ",") & vbTab
                linea &= dr("precio_15").ToString.Replace(".", ",") & vbTab
                linea &= dr("precio_16").ToString.Replace(".", ",") & vbTab
                linea &= dr("precio_17").ToString.Replace(".", ",") & vbTab
                linea &= dr("precio_18").ToString.Replace(".", ",") & vbTab
                linea &= dr("ped_0").ToString

                sw.WriteLine(linea)
            Next

            sw.Close()

        Catch ex As Exception
            da.Dispose()

        End Try

    End Sub
    Shared Sub Importar(ByVal cn As OracleConnection, ByVal Archivo As String)
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As DataTable
        Dim dr As DataRow
        Dim sr As StreamReader
        Dim linea As String
        Dim c() As String
        Dim l As Integer = 0
        Dim itm As New Articulo(cn)

        Sql = "SELECT * FROM xprecios"
        da = New OracleDataAdapter(Sql, cn)
        da.InsertCommand = New OracleCommandBuilder(da).GetInsertCommand
        da.DeleteCommand = New OracleCommandBuilder(da).GetDeleteCommand

        dt = New DataTable
        da.Fill(dt)

        sr = New StreamReader(Archivo)

        'Elimino todos los registros
        For Each dr In dt.Rows
            dr.Delete()
        Next

        'Agrego los nuevos registros
        Do Until sr.EndOfStream
            l += 1 'numero de linea del txt

            linea = sr.ReadLine
            c = Split(linea, vbTab)

            If c.GetUpperBound(0) <> 21 Then Continue Do
            'Chequeo si existe el articulo
            If Not itm.Abrir(c(0)) Then Continue Do

            dr = dt.NewRow
            dr("itmref_0") = c(0)
            dr("qty_0") = CDbl(c(1))
            dr("precio_0") = CDbl(c(2).Replace(",", "."))
            dr("precio_1") = CDbl(c(3).Replace(",", "."))
            dr("precio_2") = CDbl(c(4).Replace(",", "."))
            dr("precio_3") = CDbl(c(5).Replace(",", "."))
            dr("precio_4") = CDbl(c(6).Replace(",", "."))
            dr("precio_5") = CDbl(c(7).Replace(",", "."))
            dr("precio_6") = CDbl(c(8).Replace(",", "."))
            dr("precio_7") = CDbl(c(9).Replace(",", "."))
            dr("precio_8") = CDbl(c(10).Replace(",", "."))
            dr("precio_9") = CDbl(c(11).Replace(",", "."))
            dr("precio_10") = CDbl(c(12).Replace(",", "."))
            dr("precio_11") = CDbl(c(13).Replace(",", "."))
            dr("precio_12") = CDbl(c(14).Replace(",", "."))
            dr("precio_13") = CDbl(c(15).Replace(",", "."))
            dr("precio_14") = CDbl(c(16).Replace(",", "."))
            dr("precio_15") = CDbl(c(17).Replace(",", "."))
            dr("precio_16") = CDbl(c(18).Replace(",", "."))
            dr("precio_17") = CDbl(c(19).Replace(",", "."))
            dr("precio_18") = CDbl(c(20).Replace(",", "."))
            dr("ped_0") = CInt(c(21))
            dt.Rows.Add(dr)
        Loop

        da.Update(dt)

    End Sub

    Private Sub dtd_RowChanged(ByVal sender As Object, ByVal e As System.Data.DataRowChangeEventArgs) Handles dtd.RowChanged
        RaiseEvent PresupuestoModificado(Me)
    End Sub
    Private Sub dtd_RowDeleted(ByVal sender As Object, ByVal e As System.Data.DataRowChangeEventArgs) Handles dtd.RowDeleted
        RaiseEvent PresupuestoModificado(Me)
    End Sub
    Private Sub dtd_TableNewRow(ByVal sender As Object, ByVal e As System.Data.DataTableNewRowEventArgs) Handles dtd.TableNewRow
        RaiseEvent PresupuestoModificado(Me)
    End Sub

    Private Sub FaltanArticulosRelacionados(ByVal e As ArrayList)
        Dim arr As New DataTable
        Dim txt As String = ""
        Dim itm As Articulo
        Dim flg As Boolean = False
        Dim i As Integer 'Items faltantes

        For Each dr As DataRow In dtd.Rows
            i = 0
            If dr.RowState = DataRowState.Deleted Then Continue For
            itm = New Articulo(cn, dr("itmref_0").ToString)

            arr = ArticulosRelacionados(itm.Codigo)

            txt = "{itm} {des} requiere que exista: " & vbCrLf
            txt = txt.Replace("{itm}", itm.Codigo)
            txt = txt.Replace("{des}", itm.Descripcion)

            For Each dr2 As DataRow In arr.Rows

                If TieneArticulo(dr2("des_0").ToString) Then
                    'Alcanza con que exista unicamente este articulo
                    If CInt(dr2("req_0")) = 1 Then
                        i = 0
                        Exit For
                    End If

                Else
                    i += 1
                    txt &= "- {itm}" & vbCrLf
                    txt = txt.Replace("{itm}", dr2("des_0").ToString)

                End If

            Next

            If i > 0 Then e.Add(txt)

        Next

    End Sub
    Private Function ExisteFleteYAcarreo() As Boolean
        Dim dr As DataRow
        Dim flg As Boolean = False

        For Each dr In dtd.Rows
            If dr.RowState = DataRowState.Deleted Then Continue For

            If dr("itmref_0").ToString = ARTICULO_FLETE_Y_ACARREO Then
                Dim PrecioActual As Double = CDbl(dr("precio_0"))
                Dim PrecioSugerido As Double = CDbl(dr("precio_1"))

                flg = (PrecioActual >= (PrecioSugerido * PORCENTAJE_DESCUENTO_AUTORIZADO))
            End If
        Next

        Return flg
    End Function
    Public Function TieneIntervencion() As Boolean
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim fechaini As Date = Fecha.AddDays(-15)
        Dim fechafin As Date = Fecha.AddDays(65)

        Sql = "SELECT * "
        Sql &= "FROM interven "
        Sql &= "WHERE typ_0 IN ('B1', 'E1') AND credat_0 > :dat_0 AND credat_0 < :dat2_0 AND "
        Sql &= "      bpc_0 = :bpc_0 AND bpaadd_0 = :bpaadd_0 "
        Sql &= "ORDER BY dat_0 "

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("bpc_0", OracleType.VarChar).Value = bpc.Codigo
        da.SelectCommand.Parameters.Add("bpaadd_0", OracleType.VarChar).Value = bpa.Sucursal
        da.SelectCommand.Parameters.Add("dat_0", OracleType.DateTime).Value = fechaini
        da.SelectCommand.Parameters.Add("dat2_0", OracleType.DateTime).Value = fechafin
        da.Fill(dt)
        da.Dispose()

        Return dt.Rows.Count > 0

    End Function
    Public ReadOnly Property Presupuesto() As Presupuesto
        Get
            Return Sqh
        End Get
    End Property
    Public ReadOnly Property Pedido() As Pedido
        Get
            Return Soh
        End Get
    End Property

End Class