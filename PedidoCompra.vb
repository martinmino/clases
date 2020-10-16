Imports System.Data.OracleClient

Public Class PedidoCompra
    Implements IRuteable

    Private cn As OracleConnection
    Private dt1 As DataTable
    Private da1 As OracleDataAdapter
    Private bps As Proveedor = Nothing
    Private bpa As Sucursal = Nothing

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String

        Me.cn = cn

        Sql = "SELECT * FROM porder where pohnum_0 = :pohnum"
        da1 = New OracleDataAdapter(Sql, cn)
        da1.SelectCommand.Parameters.Add("pohnum", OracleType.VarChar)

    End Sub
    Public Function Abrir(ByVal id As String) As Boolean Implements IRuteable.Abrir
        dt1 = New DataTable
        da1.SelectCommand.Parameters("pohnum").Value = id
        da1.Fill(dt1)

        bps = Nothing
        bpa = Nothing

        Return dt1.Rows.Count > 0
    End Function
    Private Sub AbrirProveedor()
        If bps Is Nothing Then
            bps = New Proveedor(cn)
            bps.Abrir(Me.CodigoTercero)
        End If
    End Sub
    Private Sub AbrirSucursal()
        If bpa Is Nothing Then
            bpa = New Sucursal(cn)
            bpa.Abrir(Me.CodigoTercero, Me.SucursalCodigo)
        End If
    End Sub
    Public Property CarritoFecha() As Date Implements IRuteable.CarritoFecha
        Get
            Return #12/31/1599#
        End Get
        Set(ByVal value As Date)

        End Set
    End Property
    Public ReadOnly Property Tercero() As Tercero Implements IRuteable.Tercero
        Get
            AbrirProveedor()
            Return bps
        End Get
    End Property
    Public ReadOnly Property Cobranza() As Boolean Implements IRuteable.Cobranza
        Get
            Return False
        End Get
    End Property
    Public ReadOnly Property CodigoTercero() As String Implements IRuteable.CodigoTercero
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("bpsnum_0").ToString
        End Get
    End Property
    Public ReadOnly Property Domicilio() As String Implements IRuteable.Domicilio
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("bpaaddlig_0").ToString
        End Get
    End Property
    Public ReadOnly Property Equipos() As Integer Implements IRuteable.Equipos
        Get
            Return 0
        End Get
    End Property
    Public ReadOnly Property FechaEntrega() As Date Implements IRuteable.FechaEntrega
        Get
            Return #12/31/1599#
        End Get
    End Property
    Public ReadOnly Property FechaUnigis() As Date Implements IRuteable.FechaUnigis
        Get
            Return #12/31/1599#
        End Get
    End Property
    Public Property Franja1Desde() As String Implements IRuteable.Franja1Desde
        Get
            Return " "
        End Get
        Set(ByVal value As String)

        End Set
    End Property
    Public Property Franja1Hasta() As String Implements IRuteable.Franja1Hasta
        Get
            Return " "
        End Get
        Set(ByVal value As String)

        End Set
    End Property
    Public Property Franja2Desde() As String Implements IRuteable.Franja2Desde
        Get
            Return " "
        End Get
        Set(ByVal value As String)

        End Set
    End Property
    Public Property Franja2Hasta() As String Implements IRuteable.Franja2Hasta
        Get
            Return " "
        End Get
        Set(ByVal value As String)

        End Set
    End Property
    Public ReadOnly Property Hora() As String Implements IRuteable.Hora
        Get
            Return " "
        End Get
    End Property
    Public ReadOnly Property Instalaciones() As Integer Implements IRuteable.Instalaciones
        Get
            Return 0
        End Get
    End Property
    Public ReadOnly Property Localidad() As String Implements IRuteable.Localidad
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("cty_0").ToString
        End Get
    End Property
    Public ReadOnly Property Mangueras() As Integer Implements IRuteable.Mangueras
        Get
            Return 0
        End Get
    End Property
    Public Property ModoEntrega() As String Implements IRuteable.ModoEntrega
        Get
            Return " "
        End Get
        Set(ByVal value As String)

        End Set
    End Property
    Public ReadOnly Property NombreTercero() As String Implements IRuteable.NombreTercero
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("bprnam_0").ToString
        End Get
    End Property
    Public ReadOnly Property Numero() As String Implements IRuteable.Numero
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("pohnum_0").ToString
        End Get
    End Property
    Public ReadOnly Property Pedido() As Pedido Implements IRuteable.Pedido
        Get
            Return Nothing
        End Get
    End Property
    Public ReadOnly Property Peso() As Double Implements IRuteable.Peso
        Get
            Return 0
        End Get
    End Property
    Public ReadOnly Property PesoUnigis() As Double Implements IRuteable.PesoUnigis
        Get
            Return 0
        End Get
    End Property
    Public ReadOnly Property PrestamosExtintores() As Integer Implements IRuteable.PrestamosExtintores
        Get
            Return 0
        End Get
    End Property
    Public ReadOnly Property PrestamosMangueras() As Integer Implements IRuteable.PrestamosMangueras
        Get
            Return 0
        End Get
    End Property
    Public ReadOnly Property RechazosExtintor() As Integer Implements IRuteable.RechazosExtintor
        Get
            Return 0
        End Get
    End Property
    Public ReadOnly Property RechazosManguera() As Integer Implements IRuteable.RechazosManguera
        Get
            Return 0
        End Get
    End Property
    Public ReadOnly Property Remito() As String Implements IRuteable.Remito
        Get
            Return " "
        End Get
    End Property
    Public ReadOnly Property Sucursal() As Sucursal Implements IRuteable.Sucursal
        Get
            AbrirSucursal()
            Return bpa
        End Get
    End Property
    Public ReadOnly Property SucursalCodigo() As String Implements IRuteable.SucursalCodigo
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("bpaadd_0").ToString
        End Get
    End Property
    Public ReadOnly Property TieneCarro() As Boolean Implements IRuteable.TieneCarro
        Get
            Return False
        End Get
    End Property
    Public Property Tipo() As String Implements IRuteable.Tipo
        Get
            Return " "
        End Get
        Set(ByVal value As String)

        End Set
    End Property
    Public ReadOnly Property TipoTarea() As String Implements IRuteable.TipoTarea
        Get
            Return "POD"
        End Get
    End Property
    Public ReadOnly Property Varios() As Boolean Implements IRuteable.Varios
        Get
            Return True
        End Get
    End Property

End Class