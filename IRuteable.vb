Public Interface IRuteable
    ReadOnly Property Numero() As String
    ReadOnly Property TipoTarea() As String
    ReadOnly Property Equipos() As Integer
    ReadOnly Property Mangueras() As Integer
    ReadOnly Property RechazosExtintor() As Integer
    ReadOnly Property RechazosManguera() As Integer
    ReadOnly Property PrestamosExtintores() As Integer
    ReadOnly Property PrestamosMangueras() As Integer
    ReadOnly Property PesoUnigis() As Double
    ReadOnly Property Cliente() As Cliente
    ReadOnly Property Sucursal() As Sucursal
    ReadOnly Property FechaUnigis() As Date
    ReadOnly Property Pedido() As Pedido
    ReadOnly Property Remito() As String
    ReadOnly Property TieneCarro() As Boolean
    Property CarritoFecha() As Date
    Property Tipo() As String
    Property Franja1Desde() As String
    Property Franja1Hasta() As String
    Property Franja2Desde() As String
    Property Franja2Hasta() As String
End Interface