Public Interface IRuteable
    Function Abrir(ByVal id As String) As Boolean

    ReadOnly Property Numero() As String
    ReadOnly Property TipoTarea() As String
    ReadOnly Property Equipos() As Integer
    ReadOnly Property Mangueras() As Integer
    ReadOnly Property RechazosExtintor() As Integer
    ReadOnly Property RechazosManguera() As Integer
    ReadOnly Property PrestamosExtintores() As Integer
    ReadOnly Property PrestamosMangueras() As Integer
    ReadOnly Property PesoUnigis() As Double
    ReadOnly Property Tercero() As Tercero
    ReadOnly Property NombreTercero() As String
    ReadOnly Property CodigoTercero() As String
    ReadOnly Property Sucursal() As Sucursal
    ReadOnly Property SucursalCodigo() As String
    ReadOnly Property FechaUnigis() As Date
    ReadOnly Property Pedido() As Pedido
    ReadOnly Property Remito() As String
    ReadOnly Property TieneCarro() As Boolean
    ReadOnly Property Domicilio() As String
    ReadOnly Property Localidad() As String
    ReadOnly Property Instalaciones() As Integer
    ReadOnly Property Cobranza() As Boolean
    ReadOnly Property Varios() As Boolean
    ReadOnly Property Peso() As Double
    ReadOnly Property Hora() As String
    ReadOnly Property FechaEntrega() As Date
    Property ModoEntrega() As String
    Property CarritoFecha() As Date
    Property Tipo() As String
    Property Franja1Desde() As String
    Property Franja1Hasta() As String
    Property Franja2Desde() As String
    Property Franja2Hasta() As String

End Interface