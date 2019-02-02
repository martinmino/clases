Imports System.Data.OracleClient
Imports System.Windows.Forms
Imports System.IO

Public Class Contratos
    Private cn As OracleConnection
    Private dah As OracleDataAdapter 'Adaptador Tabla Cabecera
    Private dad As OracleDataAdapter 'Adaptador Tabla Detalle - sucursales
    Private dth As DataTable 'Tabla cabecera
    Private dtd As DataTable 'Tabla detalle - sucursales

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String
        Me.cn = cn

        'Adaptador de tabla cabecera
        Sql = " Select * from xcontrato where nro_0 = :nro_0"
        dah = New OracleDataAdapter(Sql, cn)
        dah.SelectCommand.Parameters.Add("nro_0", OracleType.Number)
        dah.InsertCommand = New OracleCommandBuilder(dah).GetInsertCommand
        dah.UpdateCommand = New OracleCommandBuilder(dah).GetUpdateCommand
        dah.DeleteCommand = New OracleCommandBuilder(dah).GetDeleteCommand

        'Adaptador de tabla detalle
        Sql = "SELECT * FROM  xcontratod WHERE nro_0 = :nro_0"
        dad = New OracleDataAdapter(Sql, cn)
        dad.SelectCommand.Parameters.Add("nro_0", OracleType.Number)
        dad.InsertCommand = New OracleCommandBuilder(dad).GetInsertCommand
        dad.UpdateCommand = New OracleCommandBuilder(dad).GetUpdateCommand
        dad.DeleteCommand = New OracleCommandBuilder(dad).GetDeleteCommand

        dth = New DataTable
        dtd = New DataTable
        dah.FillSchema(dth, SchemaType.Mapped)
        dad.FillSchema(dtd, SchemaType.Mapped)

    End Sub
    Public Sub Nuevo()
        Dim dr As DataRow

        dth.Clear()
        dtd.Clear()

        dr = dth.NewRow
        dr("nro_0") = 0
        dr("nombre_0") = " "
        dr("bpc_0") = " "
        dr("rep_0") = " "
        dr("bpaadd_0") = " "
        dr("dat_0") = Date.Today
        dr("duracion_0") = 12
        dr("concom_0") = " "
        dr("comtel_0") = " "
        dr("commail_0") = " "
        dr("conope_0") = " "
        dr("opetel_0") = " "
        dr("opemail_0") = " "
        dr("credat_0") = Date.Today
        dr("horam1_0") = " "
        dr("horam2_0") = " "
        dr("horat1_0") = " "
        dr("horat2_0") = " "
        dr("oc_0") = " "
        dr("phmang_0") = 0
        dr("phmangdet_0") = " "
        dr("mantbohidr_0") = 0
        dr("mantbohidd_0") = " "
        dr("mantintred_0") = 0
        dr("mantintrdd_0") = " "
        dr("m639_0") = 0
        dr("m639det_0") = " "
        dr("ext_0") = 0
        dr("extdet_0") = 0
        dr("polqui55_0") = 0
        dr("polqui90_0") = 0
        dr("polquibc_0") = 0
        dr("co2_0") = 0
        dr("espumaafff_0") = 0
        dr("halogenado_0") = 0
        dr("acetato_0") = 0
        dr("otros_0") = 0
        dr("pintura_0") = 0
        dr("repuestos_0") = 0
        dr("hidrantes_0") = 0
        dr("hidrantesd_0") = 0
        dr("tarjeta_0") = 0
        dr("visitarele_0") = 0
        dr("otrosrele_0") = 0
        dr("soporychap_0") = 0
        dr("soporycha2_0") = 0
        dr("otrosinsta_0") = 0
        dr("autoriza_0") = 0
        dr("art_0") = 0
        dr("segauto_0") = 0
        dr("estudios_0") = 0
        dr("estudiosd_0") = " "
        dr("casco_0") = 0
        dr("barbijo_0") = 0
        dr("proaudio_0") = 0
        dr("zapatos_0") = 0
        dr("otroselem_0") = 0
        dr("invmet_0") = 1
        dr("impanual_0") = 0
        dr("impmes_0") = 0
        dr("pte_0") = " "
        dr("estado_0") = 1
        dr("frecfactu_0") = 0
        dr("control_0") = 0
        dr("xunifi_0") = 0
        dr("obs_0") = " "
        dr("connum_0") = " "
        dr("srenum_0") = " "
        dth.Rows.Add(dr)

    End Sub
    Public Function Grabar() As Boolean
        Dim n As Long = Me.Numero 'Numero pedido actual
        Dim i As Integer = 0
        Dim dr As DataRow

        ' Recupero el numero del contrato actual
        dr = dth.Rows(0)

        dr.BeginEdit()
        If n = 0 Then
            n = ProximoNumero()
            dr("nro_0") = n
        End If
        dr.EndEdit()

        ' Actualizo numero ID a las sucursales
        For Each dr In dtd.Rows
            If dr.RowState = DataRowState.Deleted Then Continue For
            dr.BeginEdit()
            dr("nro_0") = n
            dr.EndEdit()
        Next

        Try
            dah.Update(dth)
            dad.Update(dtd)

        Catch ex As Exception
            Return False

        End Try

        Return True

    End Function
    Public Function Abrir(ByVal Nro As Long) As Boolean
        dth.Clear()
        dtd.Clear()

        dah.SelectCommand.Parameters("nro_0").Value = Nro
        dad.SelectCommand.Parameters("nro_0").Value = Nro

        dah.Fill(dth)
        dad.Fill(dtd)

        Return dth.Rows.Count > 0

    End Function
    Private Function ProximoNumero() As Integer
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim Nro As Integer = 0

        da = New OracleDataAdapter("SELECT MAX(nro_0) FROM xcontrato", cn)
        da.Fill(dt)
        da.Dispose()

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            If Not IsDBNull(dr(0)) Then Nro = CInt(dr(0))
        End If

        Nro += 1

        Return Nro

    End Function
    Public ReadOnly Property Sucursales() As DataTable
        Get
            Return dtd
        End Get
    End Property
    Public ReadOnly Property Numero() As Integer
        Get
            Dim dr As DataRow
            Dim n As Integer = 0

            If dth.Rows.Count > 0 Then
                dr = dth.Rows(0)
                n = CInt(dr("nro_0"))
            End If

            Return n
        End Get
    End Property
    Public Property Nombre() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("nombre_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("nombre_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property Cliente() As Cliente
        Get
            Dim dr As DataRow
            Dim bpc As New Cliente(cn)

            dr = dth.Rows(0)
            bpc.Abrir(dr("bpc_0").ToString)

            Return bpc

        End Get

    End Property
    Public Property ClienteCodigo() As String
        Get
            Dim dr As DataRow

            dr = dth.Rows(0)

            Return dr("bpc_0").ToString

        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            dr = dth.Rows(0)

            dr.BeginEdit()
            dr("bpc_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Sucursal() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("bpaadd_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("bpaadd_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property FechaCreacion() As Date
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CDate(dr("credat_0"))
        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("credat_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property FechaInicio() As Date
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CDate(dr("dat_0"))
        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("dat_0") = New Date(value.Year, value.Month, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Duracion() As Integer
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("duracion_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("duracion_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property ContactoComercial() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("concom_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("concom_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property ContactoComercialTelefono() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("comtel_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("comtel_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property ContactoComercialMail() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("commail_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("commail_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property ContactoOperador() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("conope_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("conope_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property ContactoOperadorTelefono() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("opetel_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("opetel_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property ContactoOperadorMail() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("opemail_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("opemail_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property TurnoMananaDesde() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("horam1_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("horam1_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property TurnoMananahasta() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("horam2_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("horam2_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property TurnoTardeDesde() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("horat1_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("horat1_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property TurnoTardehasta() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("horat2_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("horat2_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Vendedor() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("rep_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("rep_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property OC() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("oc_0").ToString.Trim
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("oc_0") = If(value.Trim = "", " ", value)
            dr.EndEdit()
        End Set
    End Property
    Public Property Observaciones() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("obs_0").ToString.Trim
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("obs_0") = If(value.Trim = "", " ", value)
            dr.EndEdit()
        End Set
    End Property
    Public Property PhMagueras() As Boolean
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("phmang_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("phmang_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property PhManguerasDetalle() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("PHMANGDET_0").ToString.Trim
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("PHMANGDET_0") = If(value.Trim = "", " ", value)
            dr.EndEdit()
        End Set
    End Property
    Public Property MantBocaHidrantes() As Boolean
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("MANTBOHIDR_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("MANTBOHIDR_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property MantBocaHidrantesDetalle() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("MANTBOHIDD_0").ToString.Trim
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("MANTBOHIDD_0") = If(value.Trim = "", " ", value)
            dr.EndEdit()
        End Set
    End Property
    Public Property MantIntegralRedHidrantes() As Boolean
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("MANTINTRED_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("MANTINTRED_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property MantIntegralRedHidrantesDetalle() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("MANTINTRDD_0").ToString.Trim
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("MANTINTRDD_0") = If(value.Trim = "", " ", value)
            dr.EndEdit()
        End Set
    End Property
    Public Property M639() As Boolean
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("m639_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("m639_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property M639Detalle() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("m639det_0").ToString.Trim
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("m639det_0") = If(value.Trim = "", " ", value)
            dr.EndEdit()
        End Set
    End Property
    Public Property TieneMatafuegos() As Integer
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("ext_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("ext_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property CantidadMatafuegos() As Integer
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("EXTDET_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("EXTDET_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property PolvoQuimicoABC55() As Boolean
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("POLQUI55_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("POLQUI55_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property PolvoQuimicoABC90() As Boolean
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("POLQUI90_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("POLQUI90_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property PolvoQuimicoBC() As Boolean
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("POLQUIBC_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("POLQUIBC_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property CO2() As Boolean
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("co2_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("co2_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property EspumaAFFF() As Boolean
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("ESPUMAAFFF_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("ESPUMAAFFF_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property HALOGENADO() As Boolean
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("HALOGENADO_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("HALOGENADO_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property ACETATO() As Boolean
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("ACETATO_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("ACETATO_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property OtrosMatafuegos() As Boolean
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("otros_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("otros_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Pintura() As Integer
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("PINTURA_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("PINTURA_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Repuestos() As Integer
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("REPUESTOS_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("REPUESTOS_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Hidrantes() As Boolean
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("HIDRANTES_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("HIDRANTES_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property CantidadHidrantes() As Integer
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("HIDRANTESD_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("HIDRANTESD_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Tarjeta() As Integer
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("TARJETA_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("TARJETA_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property VisitaRelevamiento() As Boolean
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("VISITARELE_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("VISITARELE_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property ControlPeriodico() As Boolean
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("CONTROL_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("CONTROL_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property OtrosRelevamientos() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("OTROSRELE_0").ToString.Trim
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("OTROSRELE_0") = If(value.Trim = "", " ", value)
            dr.EndEdit()
        End Set
    End Property
    Public Property SoporteChapaBalizaProvision() As Boolean
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("SOPORYCHAP_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("SOPORYCHAP_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property SoporteChapaBalizaInstalacion() As Boolean
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("SOPORYCHA2_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("SOPORYCHA2_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property OtrasInstalacion() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("OTROSINSTA_0").ToString.Trim
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("OTROSINSTA_0") = If(value.Trim = "", " ", value)
            dr.EndEdit()
        End Set
    End Property
    Public Property AutorizacionPersonal() As Boolean
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("AUTORIZA_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("AUTORIZA_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property ComprobantePagoART() As Boolean
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("ART_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("art_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property SeguroVehiculo() As Boolean
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("SEGAUTO_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("segauto_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property EstudiosMedicos() As Integer
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("ESTUDIOS_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("ESTUDIOS_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property EstudiosMedicosDetalle() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("ESTUDIOSD_0").ToString.Trim
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("ESTUDIOSD_0") = If(value.Trim = "", " ", value)
            dr.EndEdit()
        End Set
    End Property
    Public Property Casco() As Boolean
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("casco_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("casco_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Barbijo() As Boolean
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("barbijo_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("barbijo_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property ProtectorAuditivo() As Boolean
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("PROAUDIO_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("PROAUDIO_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property ZapatosDeSeguridad() As Boolean
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("ZAPATOS_0")) = 2
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("ZAPATOS_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property OtrosElementos() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("OTROSELEM_0").ToString.Trim
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("OTROSELEM_0") = If(value.Trim = "", " ", value)
            dr.EndEdit()
        End Set
    End Property
    Public Property Unificacion() As Integer
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("xunifi_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("xunifi_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property FrecuenciaFacturacion() As Integer
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("FRECFACTU_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("FRECFACTU_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property ImporteAnual() As Long
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CLng(dr("IMPANUAL_0"))
        End Get
        Set(ByVal value As Long)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("IMPANUAL_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property ImporteMensual() As Long
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CLng(dr("IMPMES_0"))
        End Get
        Set(ByVal value As Long)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("IMPMES_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property CondicionPago() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("PTE_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("PTE_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Estado() As Integer
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("estado_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("estado_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property ModoFacturacion() As Integer
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("invmet_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("invmet_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Solicitud() As String
        Get
            Dim txt As String = ""
            Dim dr As DataRow

            If dth.Rows.Count > 0 Then
                dr = dth.Rows(0)
                txt = dr("srenum_0").ToString
            End If

            Return txt
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dth.Rows.Count > 0 Then
                dr = dth.Rows(0)
                dr.BeginEdit()
                dr("srenum_0") = value
                dr.EndEdit()
            End If
            
        End Set
    End Property
    Public Sub SucursalesCubiertas(ByVal dt As DataTable)
        For Each dr As DataRow In dtd.Rows
            dr.Delete()
        Next

        For Each dr As DataRow In dt.Rows
            Dim dr1 As DataRow = dtd.NewRow

            dr1("nro_0") = Numero
            dr1("bpaadd_0") = dr("bpaadd_0").ToString

            dtd.Rows.Add(dr1)
        Next

    End Sub
    Public Property ContratoAdonix() As String
        Get
            Dim txt As String = ""

            If dth.Rows.Count > 0 Then
                Dim dr As DataRow = dtd.Rows(0)
                txt = dr("connum_0").ToString
            End If

            Return txt

        End Get
        Set(ByVal value As String)
            If dth.Rows.Count > 0 Then
                Dim dr As DataRow = dth.Rows(0)
                dr.BeginEdit()
                dr("connum_0") = IIf(value.Trim = "", " ", value.Trim)
                dr.EndEdit()

            End If
        End Set
    End Property

End Class
