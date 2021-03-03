Public Class GeneradorQrAfip
    'Version del formato de los datos segun sistema de AFIP
    Private Const VERSION_AFIP As Integer = 1

    Public FechaEmision As Date
    Public CuitEmitor As Long
    Public PuntoVenta As Integer
    Public TipoComprobante As Integer
    Public NumeroComprobante As Integer
    Public ImporteTotal As Double
    Public Moneda As String
    Public Cotizacion As Double
    Public TipoDocumentoReceptor As Integer
    Public NumeroDocumentoReceptor As Long
    Public CodigoTipoAutorizacion As String
    Public CodigoAutorizacion As Long

    Private _Qr As String
    Private _JSon As String

    Public Sub Nuevo()
        FechaEmision = Nothing
        CuitEmitor = 0
        PuntoVenta = 0
        TipoComprobante = 0
        NumeroComprobante = 0
        ImporteTotal = 0
        Moneda = ""
        Cotizacion = 0
        TipoDocumentoReceptor = 0
        NumeroDocumentoReceptor = 0
        CodigoTipoAutorizacion = ""
        CodigoAutorizacion = 0

    End Sub
    Public Sub GenerarQr()

        _JSon = "{"
        _JSon &= "'ver':" & VERSION_AFIP.ToString & ","
        _JSon &= "'fecha':'" & FechaEmision.ToString("yyyy-MM-dd") & "',"
        _JSon &= "'cuit':" & CuitEmitor & ","
        _JSon &= "'ptoVta':" & PuntoVenta.ToString & ","
        _JSon &= "'tipoCmp':" & TipoComprobante.ToString & ","
        _JSon &= "'nroCmp':" & NumeroComprobante.ToString & ","
        _JSon &= "'importe':" & FormatearImporte(ImporteTotal) & ","
        _JSon &= "'moneda':'" & Moneda & "',"
        _JSon &= "'ctz':" & FormatearImporte(Cotizacion) & ","
        If TipoDocumentoReceptor > 0 Then
            _JSon &= "'tipoDocRec':" & TipoDocumentoReceptor & ","
        End If
        If NumeroDocumentoReceptor.ToString.Trim <> "" Then
            _JSon &= "'nroDocRec':" & NumeroDocumentoReceptor & ","
        End If
        _JSon &= "'tipoCodAut':'" & CodigoTipoAutorizacion & "',"
        _JSon &= "'codAut':" & CodigoAutorizacion.ToString
        _JSon &= "}"

        _JSon = _JSon.Replace("'", Chr(34))

        Dim a() As Byte = System.Text.Encoding.UTF8.GetBytes(_JSon)

        _Qr = System.Convert.ToBase64String(a)

    End Sub
    Private Function FormatearImporte(ByVal v As Double) As String
        'Obtengo la parte entera
        Dim ParteEntera As Double
        Dim ParteDecimal As Double
        Dim Salida As String

        ParteEntera = Fix(v)
        ParteDecimal = v - ParteEntera
        ParteDecimal *= 100
        ParteDecimal = Fix(ParteDecimal)

        Salida = ParteEntera.ToString
        If ParteDecimal > 0 Then
            Salida &= "." & ParteDecimal.ToString
        End If

        Return Salida

    End Function
    Public ReadOnly Property Qr() As String
        Get
            Return Me._Qr
        End Get
    End Property
    Public ReadOnly Property JSon() As String
        Get
            Return _JSon
        End Get
    End Property


End Class
