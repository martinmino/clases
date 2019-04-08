Imports System.Data.OracleClient
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine

Public Class Remito
    Implements IRuteable

    Private cn As OracleConnection
    Private da1 As OracleDataAdapter
    Private dt1 As New DataTable
    Private da2 As OracleDataAdapter
    Private dt2 As New DataTable
    Private da3 As OracleDataAdapter
    Private dt3 As DataTable
    Public Const DB_USR As String = "GEOPROD"
    Public Const DB_PWD As String = "tiger"
    Private bpa As Sucursal = Nothing

    Private l_Equipos As Integer
    Private l_Mangas As Integer
    Private l_PrestamosExt As Integer
    Private l_PrestamosMan As Integer
    Private l_Instalaciones As Integer
    Private l_RechazosExt As Integer
    Private l_RechazosMan As Integer
    Private l_Varios As Boolean
    Private l_Peso As Double
    Private l_Peso2 As Double
    Private l_Hora As String
    Private l_TieneInstalacion As Boolean
    Private l_EsTarea As Boolean
    Private l_Serie As String

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn

        Adaptadores()

    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal Numero As String)
        Me.New(cn)
    End Sub
    Private Sub Adaptadores()
        Dim Sql As String

        Sql = "SELECT * FROM sdelivery WHERE sdhnum_0 = :sdhnum_0"
        da1 = New OracleDataAdapter(Sql, cn)
        da1.SelectCommand.Parameters.Add("sdhnum_0", OracleType.VarChar)
        da1.UpdateCommand = New OracleCommandBuilder(da1).GetUpdateCommand

        Sql = "SELECT * FROM sdeliveryd WHERE sdhnum_0 = :sdhnum_0"
        da2 = New OracleDataAdapter(Sql, cn)
        da2.SelectCommand.Parameters.Add("sdhnum_0", OracleType.VarChar)
        da2.UpdateCommand = New OracleCommandBuilder(da2).GetUpdateCommand

        Sql = "SELECT gac.* "
        Sql &= "FROM sinvoicev sih INNER JOIN gaccdudate gac ON (sih.num_0 = gac.num_0) "
        Sql &= "WHERE sihorinum_0 = :num"
        da3 = New OracleDataAdapter(Sql, cn)
        da3.SelectCommand.Parameters.Add("num", OracleType.VarChar)

    End Sub
    Public Function Abrir(ByVal Nro As String) As Boolean
        dt1.Clear()
        dt2.Clear()
        dt3 = Nothing

        bpa = Nothing

        l_Equipos = 0
        l_Mangas = 0
        l_PrestamosExt = 0
        l_PrestamosMan = 0
        l_Instalaciones = 0
        l_RechazosExt = 0
        l_RechazosMan = 0
        l_Varios = False
        l_Peso = 0
        l_Peso2 = 0
        l_Hora = ""
        l_TieneInstalacion = False
        l_EsTarea = False

        da1.SelectCommand.Parameters("sdhnum_0").Value = Nro
        da1.Fill(dt1)

        da2.SelectCommand.Parameters("sdhnum_0").Value = Nro
        da2.Fill(dt2)

        AnalizarRemito()

        Return dt1.Rows.Count > 0
    End Function
    Public Sub Grabar()
        da1.Update(dt1)
    End Sub
    Public Function EtiquetasEntrega(ByVal Archivo As String, ByVal Imprimir As Boolean) As Boolean
        Select Case Me.Cliente.Codigo
            Case "300006" ' INC
                EtiquetasINC(Archivo, Imprimir)

            Case Else
                EtiquetasGeneral(Archivo, Imprimir)

        End Select
    End Function
    Public Function EtiquetasINC(ByVal Archivo As String, ByVal Imprimir As Boolean) As Boolean
        Dim itmbpc As New ArticuloCliente(cn)
        Dim bpc As Cliente = Me.Cliente 'Cliente del remito
        Dim bpa As Sucursal = bpc.Sucursal(Me.SucursalCodigo) ' Sucursal del remito
        Dim soh As Pedido = Me.Pedido 'Pedido del remito
        Dim Ped As String = " " ' Nro Pedido cliente
        Dim z() As String
        'Puerto de impresora destino
        Dim prn As New Impresora(cn, "BULTOS")

        'Array para obtener pedido cliente de la referencia del pedido nuestro
        z = Split(soh.Referencia, "/")

        If z.Count = 2 Then Ped = z(1)

        Dim TotalBultos As Integer = 0 ' Total de bultos
        Dim CantidadExtintores As Integer ' Cantidad
        Dim txt As String = ""
        Dim st As Stream
        Dim sw As StreamWriter

        For Each dr As DataRow In Detalle.Rows
            If itmbpc.Abrir(dr("itmref_0").ToString, bpc.Codigo) Then

                CantidadExtintores = CInt(dr("qty_0"))

                'Obtengo la cantidad de bultos
                TotalBultos += CantidadExtintores \ itmbpc.CantidadPorCaja

                'Si la divicion no es entera sumo un bulto mas
                If CantidadExtintores Mod itmbpc.CantidadPorCaja > 0 Then TotalBultos += 1
            End If
        Next

        'Cabecera del archivo ZEBRA
        txt = "I8,A,001" & vbCrLf
        txt &= "Q328,024 " & vbCrLf
        txt &= "q831" & vbCrLf
        txt &= "rN" & vbCrLf
        txt &= "S4" & vbCrLf
        txt &= "D7" & vbCrLf
        txt &= "ZT" & vbCrLf
        txt &= "JF" & vbCrLf

        'DETALLE DEL ARCHIVO ZEBRA
        For i As Integer = 1 To TotalBultos
            txt &= "OD" & vbCrLf
            txt &= "R191,0" & vbCrLf
            txt &= "N" & vbCrLf
            txt &= "A5,9,0,4,1,1,N,""NRO. OC: {OC}""" & vbCrLf
            txt &= "A5,35,0,4,1,1,N,""NRO. REMITO: {REMITO}""" & vbCrLf
            txt &= "A5,81,0,4,1,1,N,""INC.SA""" & vbCrLf
            'txt &= "A5,108,0,4,1,1,N,""{NOM_CLIENTE}""" & vbCrLf
            txt &= "A5,154,0,4,1,1,N,""TIENDA:""" & vbCrLf
            txt &= "A5,183,0,4,1,1,N,""{TIENDA}""" & vbCrLf
            txt &= "A5,246,0,4,2,2,N,""BULTO: {BULTO_NRO}/{BULTO_TOTAL}""" & vbCrLf
            txt &= "A5,310,0,4,1,1,N,""{ARTICULO}""" & vbCrLf
            txt &= "P2" & vbCrLf

            txt = txt.Replace("{OC}", soh.Referencia)
            txt = txt.Replace("{REMITO}", NumeroFormateado)
            txt = txt.Replace("{TIENDA}", bpa.Direccion2)
            txt = txt.Replace("{BULTO_NRO}", i.ToString)
            txt = txt.Replace("{BULTO_TOTAL}", TotalBultos.ToString)
            txt = txt.Replace("{ARTICULO}", "MATAFUEGOS DONNY SRL")
        Next


        'Grabo archivo
        st = File.Open(Archivo, FileMode.Create, FileAccess.Write, FileShare.None)
        sw = New StreamWriter(st)
        sw.Write(txt)
        sw.Close()
        st.Close()

        If Imprimir Then File.Copy(Archivo, prn.RecursoRed)

    End Function
    Public Function EtiquetasGeneral(ByVal Archivo As String, ByVal Imprimir As Boolean) As Boolean
        Dim itmbpc As New ArticuloCliente(cn)
        Dim bpc As Cliente = Me.Cliente 'Cliente del remito
        Dim bpa As Sucursal = bpc.Sucursal(Me.SucursalCodigo) ' Sucursal del remito
        Dim soh As Pedido = Me.Pedido 'Pedido del remito
        Dim Ped As String = " " ' Nro Pedido cliente
        Dim z() As String
        'Puerto de impresora destino
        Dim prn As New Impresora(cn, "BULTOS")

        'Array para obtener pedido cliente de la referencia del pedido nuestro
        z = Split(soh.Referencia, "/")

        If z.Count = 2 Then Ped = z(1)

        Dim TotalBultos As Integer = 0 ' Total de bultos
        Dim CantidadExtintores As Integer ' Cantidad
        Dim txt As String = ""
        Dim st As Stream
        Dim sw As StreamWriter

        For Each dr As DataRow In Detalle.Rows
            If itmbpc.Abrir(dr("itmref_0").ToString, bpc.Codigo) Then

                CantidadExtintores = CInt(dr("qty_0"))

                'Obtengo la cantidad de bultos
                TotalBultos = CantidadExtintores \ itmbpc.CantidadPorCaja

                'Si la divicion no es entera sumo un bulto mas
                If CantidadExtintores Mod itmbpc.CantidadPorCaja > 0 Then TotalBultos += 1

                'Cabecera del archivo ZEBRA
                txt = "I8,A,001" & vbCrLf
                txt &= "Q328,024 " & vbCrLf
                txt &= "q831" & vbCrLf
                txt &= "rN" & vbCrLf
                txt &= "S4" & vbCrLf
                txt &= "D7" & vbCrLf
                txt &= "ZT" & vbCrLf
                txt &= "JF" & vbCrLf

                'DETALLE DEL ARCHIVO ZEBRA
                For i As Integer = 1 To TotalBultos
                    txt &= "OD" & vbCrLf
                    txt &= "R191,0" & vbCrLf
                    txt &= "N" & vbCrLf
                    txt &= "A5,9,0,4,1,1,N,""NRO. PEDIDO: {PEDIDO}""" & vbCrLf
                    txt &= "A5,35,0,4,1,1,N,""NRO. REMITO: {REMITO}""" & vbCrLf
                    txt &= "A5,81,0,4,1,1,N,""{COD_CLIENTE}""" & vbCrLf
                    txt &= "A5,108,0,4,1,1,N,""{NOM_CLIENTE}""" & vbCrLf
                    txt &= "A5,154,0,4,1,1,N,""{DOMICILIO}""" & vbCrLf
                    txt &= "A5,183,0,4,1,1,N,""{CIUDAD}""" & vbCrLf
                    txt &= "A5,246,0,4,2,2,N,""BULTO: {BULTO_NRO}/{BULTO_TOTAL}""" & vbCrLf
                    txt &= "A5,310,0,4,1,1,N,""{ARTICULO}""" & vbCrLf
                    txt &= "P2" & vbCrLf

                    txt = txt.Replace("{PEDIDO}", Ped)
                    txt = txt.Replace("{REMITO}", NumeroFormateado)
                    txt = txt.Replace("{COD_CLIENTE}", bpa.Nombre)
                    txt = txt.Replace("{NOM_CLIENTE}", bpa.Direccion3)
                    txt = txt.Replace("{DOMICILIO}", bpa.Direccion)
                    txt = txt.Replace("{CIUDAD}", bpa.CodigoPostal & " - " & bpa.Ciudad)
                    txt = txt.Replace("{BULTO_NRO}", i.ToString)
                    txt = txt.Replace("{BULTO_TOTAL}", TotalBultos.ToString)
                    txt = txt.Replace("{ARTICULO}", dr("itmdes1_0").ToString)
                Next

                'Grabo archivo
                st = File.Open(Archivo, FileMode.Create, FileAccess.Write, FileShare.None)
                sw = New StreamWriter(st)
                sw.Write(txt)
                sw.Close()
                st.Close()

                If Imprimir Then File.Copy(Archivo, prn.RecursoRed)
            End If
        Next

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

    'PROPERTY
    Public ReadOnly Property FechaUnigis() As Date Implements IRuteable.FechaUnigis
        Get
            'Fecha de pistoleo a Logistica o Abonos
            Dim Sectores() As String = {"ABO", "LOG"}
            Dim f As Date = Nothing
            Dim p As New Seguimiento(cn)
            p.Abrir(Me)

            f = p.UltimaFechaEnviadoA(Sectores)

            Return f

        End Get
    End Property
    Public ReadOnly Property TieneDatos() As Boolean
        Get
            Return dt1.Rows.Count > 0
        End Get
    End Property
    Public ReadOnly Property Numero() As String Implements IRuteable.Numero
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("sdhnum_0").ToString
        End Get
    End Property
    Public ReadOnly Property NumeroFormateado() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)

            Dim suc As String = dr("sdhnum_0").ToString.Substring(6, 4)
            Dim Num As String = dr("sdhnum_0").ToString.Substring(10)

            Return suc & "-" & Num

        End Get
    End Property
    Public ReadOnly Property PedidoCodigo() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("sohnum_0").ToString
        End Get
    End Property
    Public ReadOnly Property Cliente() As Cliente Implements IRuteable.Cliente
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Dim bpc As New Cliente(cn)
            bpc.Abrir(dr("bpcord_0").ToString)
            Return bpc
        End Get
    End Property
    Public ReadOnly Property Pedido() As Pedido Implements IRuteable.Pedido
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Dim soh As New Pedido(cn)
            soh.Abrir(dr("sohnum_0").ToString)
            Return soh
        End Get
    End Property
    Public ReadOnly Property Sucursal() As Sucursal Implements IRuteable.Sucursal
        Get
            If bpa Is Nothing Then
                Dim dr As DataRow = dt1.Rows(0)
                bpa = New Sucursal(cn, Me.Cliente.Codigo, dr("bpaadd_0").ToString)
            End If

            Return bpa
        End Get
    End Property
    Public ReadOnly Property SucursalCodigo() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("bpaadd_0").ToString
        End Get
    End Property
    Public ReadOnly Property SucursalCalle() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("bpdaddlig_0").ToString
        End Get
    End Property
    Public ReadOnly Property Facturado() As Boolean
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return (CInt(dr("invflg_0")) = 2)
        End Get
    End Property
    Public ReadOnly Property Factura() As String
        Get
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                Return dr("sihnum_0").ToString()

            Else
                Return ""

            End If

        End Get

    End Property
    Public ReadOnly Property ModoEntrega() As String
        Get
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                Return dr("mdl_0").ToString()

            Else
                Return ""

            End If

        End Get

    End Property
    Public Property Ruta() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("xruta_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt1.Rows(0)
            dr.BeginEdit()
            dr("xruta_0") = value
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
            dr.BeginEdit()
            dr("xcer_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Despachado() As Boolean
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return (CInt(dr("xsalio_0")) = 2)
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow = dt1.Rows(0)
            dr.BeginEdit()
            dr("xsalio_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Estado() As Integer
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return CInt(dr("xflgrto_0"))

        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow = dt1.Rows(0)
            dr.BeginEdit()
            dr("xflgrto_0") = value
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
    Public Property Observaciones() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("xobserva_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt1.Rows(0)
            dr.BeginEdit()
            dr("xobserva_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property RemitoDevuelto() As Boolean
        Get
            Dim Sql As String
            Dim da As OracleDataAdapter
            Dim dt As New DataTable
            Dim dr As DataRow
            Dim flg As Boolean = False

            Sql = "SELECT SUM(qty_0), SUM(rtnqty_0) "
            Sql &= "FROM sdeliveryd "
            Sql &= "WHERE sdhnum_0 = :sdhnum_0 "
            da = New OracleDataAdapter(Sql, cn)
            da.SelectCommand.Parameters.Add("sdhnum_0", OracleType.VarChar).Value = Numero
            da.Fill(dt)

            dr = dt.Rows(0)

            If Not (IsDBNull(dr(0)) OrElse IsDBNull(dr(1))) Then
                flg = CInt(dr(1)) >= CInt(dr(0))
            End If

            da.Dispose() : da = Nothing
            dt.Dispose() : dt = Nothing

            Return flg

        End Get
    End Property
    Public ReadOnly Property MueveStock() As Boolean
        Get
            Dim Sql As String = "SELECT * FROM sdeliveryd WHERE stomgtcod_0 <> 1 AND sdhnum_0 = :sdhnum_0"
            Dim da As OracleDataAdapter
            Dim dtx As New DataTable

            da = New OracleDataAdapter(Sql, cn)
            da.SelectCommand.Parameters.Add("sdhnum_0", OracleType.VarChar).Value = Numero

            da.Fill(dtx)
            da.Dispose()

            Return (dtx.Rows.Count > 0)

        End Get
    End Property
    Public ReadOnly Property Detalle() As DataTable
        Get
            Dim dt As New DataTable
            Dim da As OracleDataAdapter
            Dim Sql As String

            Sql = "SELECT * FROM sdeliveryd WHERE sdhnum_0 = :sdhnum"
            da = New OracleDataAdapter(Sql, cn)
            da.SelectCommand.Parameters.Add("sdhnum", OracleType.VarChar).Value = Numero
            da.Fill(dt)
            da.Dispose()

            Return dt
        End Get
    End Property
    Public ReadOnly Property Seguimiento() As Seguimiento
        Get
            Dim segto As New Seguimiento(cn)
            segto.Abrir(Me)
            Return segto
        End Get
    End Property
    Public ReadOnly Property TipoTarea() As String Implements IRuteable.TipoTarea
        Get
            If l_TieneInstalacion Then
                If l_Equipos > 0 Or l_Varios Then
                    Return "NCI"

                Else
                    Return "INS"

                End If
            Else
                Return "NUE"

            End If

        End Get
    End Property
    Public ReadOnly Property Equipos() As Integer Implements IRuteable.Equipos
        Get
            Return l_Equipos
        End Get
    End Property
    Public ReadOnly Property Mangueras() As Integer Implements IRuteable.Mangueras
        Get
            Return l_Mangas
        End Get
    End Property
    Public ReadOnly Property RechazosExtintor() As Integer Implements IRuteable.RechazosExtintor
        Get
            Return l_RechazosExt
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
    Public Property Franja1Desde() As String Implements IRuteable.Franja1Desde
        Get
            Return Sucursal.TurnoMananaDesde
        End Get
        Set(ByVal value As String)
        End Set
    End Property
    Public Property Franja2Desde() As String Implements IRuteable.Franja2Desde
        Get
            Return Sucursal.TurnoTardeDesde
        End Get
        Set(ByVal value As String)
        End Set
    End Property
    Public Property Franja1Hasta() As String Implements IRuteable.Franja1Hasta
        Get
            Return Sucursal.TurnoMananaHasta
        End Get
        Set(ByVal value As String)
        End Set
    End Property
    Public Property Franja2Hasta() As String Implements IRuteable.Franja2Hasta
        Get
            Return Sucursal.TurnoTardeHasta
        End Get
        Set(ByVal value As String)
        End Set
    End Property
    Private Sub AnalizarRemito()
        Dim dr As DataRow
        Dim itm As New Articulo(cn)

        'Rutina de análisis de Remitos de nuevos
        For Each dr In dt2.Rows
            Dim Qty As Integer = CInt(dr("qty_0"))

            itm.Abrir(dr("itmref_0").ToString)

            If itm.Categoria = "10" Then
                l_Equipos += Qty
                l_Peso += Qty * itm.peso
                l_Peso2 += Qty * itm.peso

            ElseIf itm.LineaProducto = "651" Then
                l_TieneInstalacion = True
                l_Instalaciones += Qty

            Else
                l_Peso += Qty * itm.peso
                l_Peso2 += Qty * itm.peso

                l_Varios = True

            End If

        Next

    End Sub
    Public ReadOnly Property TieneCarro() As Boolean Implements IRuteable.TieneCarro
        Get
            Dim Sql As String
            Dim da As OracleDataAdapter
            Dim dt As New DataTable

            Sql = "SELECT DISTINCT sdd.itmref_0 "
            Sql &= "FROM sdeliveryd sdd INNER JOIN itmmaster itm ON sdd.itmref_0 = itm.itmref_0 "
            Sql &= "WHERE xcarro_0 = 2 AND "
            Sql &= "	  sdhnum_0 = :sdhnum"

            da = New OracleDataAdapter(Sql, cn)
            da.SelectCommand.Parameters.Add("sdhnum", OracleType.VarChar).Value = Me.Numero
            Try
                da.Fill(dt)
                da.Dispose()

            Catch ex As Exception
                Return False
            End Try

            Return dt.Rows.Count > 0
        End Get
    End Property
    Public Property Tipo() As String Implements IRuteable.Tipo
        Get
            Return ""
        End Get
        Set(ByVal value As String)

        End Set
    End Property
    Public ReadOnly Property Remito() As String Implements IRuteable.Remito
        Get
            Return Numero
        End Get
    End Property
    Public Property CarritoFecha() As Date Implements IRuteable.CarritoFecha
        Get
            Return #12/31/1599#
        End Get
        Set(ByVal value As Date)

        End Set
    End Property
    Public ReadOnly Property TieneDevolucion() As Boolean
        Get
            Dim Sql As String = "select * from sdeliveryd where sdhnum_0 = :sdhnum_0 and rtnqty_0 > 0"
            Dim da As OracleDataAdapter
            Dim dtx As New DataTable

            da = New OracleDataAdapter(Sql, cn)
            da.SelectCommand.Parameters.Add("sdhnum_0", OracleType.VarChar).Value = Numero

            da.Fill(dtx)
            da.Dispose()

            Return (dtx.Rows.Count > 0)
        End Get
    End Property
    Public ReadOnly Property PedidoNumero() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Dim soh As New Pedido(cn)
            Return dr("sohnum_0").ToString
        End Get
    End Property
    Public Sub EnvioMailAvisoMostrador()
        Dim txt As String = ""
        Dim eMail As New CorreoElectronico
        Dim sih As New Factura(cn)
        Dim rep As New Vendedor(cn)

        'Salgo si no encuentro la factura
        If Not sih.Abrir(Me.Factura) Then
            Exit Sub
        End If
        If sih.Sociedad.Codigo <> "DNY" Then Exit Sub
        rep = Me.Cliente.Vendedor 'Obtengo vendedor

        If Me.Pedido.Referencia = "H" Then
            txt &= "<p>" & rep.Nombre & "</p> "
            txt &= "<p>Le informamos que el cliente " & Me.Cliente.Nombre & ", tiene un pedido (" & Me.Pedido.Numero & ") listo para retirar </p> "
            Try
                With eMail
                    .Remitente("noreply@matafuegosgeorgia.com")
                    .AgregarDestinatario(rep.Mail)
                    .Asunto = "Aviso de pedido listo"
                    .EsHtml = True
                    .Cuerpo = txt
                    If .CantidadTo > 0 Then .Enviar(True)
                End With

            Catch ex As Exception
            End Try
        Else
            txt &= "<p>" & Me.Cliente.Nombre & "</p> "
            txt &= "<p>Me comunico desde Matafuegos Georgia para notificarle que el pedido esta listo para ser retirado<strong>, "
            txt &= "</strong>para lo cual deber&aacute; anunciarse en Manuel A. Rodr&iacute;guez 2838, donde le har&aacute;n entrega "
            txt &= "de su documentaci&oacute;n para&nbsp;retirar por nuestro dep&oacute;sito de lunes a viernes, en el&nbsp; "
            txt &= "<strong>horario de 9 a 12.30 y de 14 a 16.30 hs</strong>.</p> "
            txt &= "<p>Cordialmente,</p> "
            txt &= "<p>{vendedor}</p> "
            txt &= "<p>Matafuegos Georgia</p> "

            ' rep = Me.Cliente.Vendedor 'Obtengo vendedor

            'Reemplazo de marcas
            'txt = txt.Replace("{fecha}", Me..ToString("dd/MM/yyyy"))
            txt = txt.Replace("{itn}", Me.Numero)
            If rep.Codigo = "17" Then
                txt = txt.Replace("{vendedor}", rep.Analista.Nombre.ToUpper)
                txt = txt.Replace("{interno}", "")
            Else
                txt = txt.Replace("{vendedor}", rep.Nombre.ToUpper)
                txt = txt.Replace("{interno}", rep.Interno)
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
                    '.AgregarDestinatario(Me.Cliente.MailFC)
                    '.AgregarDestinatarioCopia(rep.Mail)
                    .AgregarDestinatario("ioeyen@matafuegosgeorgia.com")
                    .Asunto = "Aviso de pedido listo"
                    .EsHtml = True
                    .Cuerpo = txt
                    .AdjuntarArchivo(sih.Numero & ".pdf")
                    If .CantidadTo > 0 Then .Enviar(True)
                End With

            Catch ex As Exception
            End Try
            Try
                File.Delete(sih.Numero & ".pdf")

            Catch ex As Exception
            End Try
        End If
    End Sub
End Class
