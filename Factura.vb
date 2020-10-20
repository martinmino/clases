Imports System.Data.OracleClient

Public Class Factura

    Private cn As OracleConnection
    Private da1 As OracleDataAdapter
    Private da2 As OracleDataAdapter
    Private da3 As OracleDataAdapter
    Private da4 As OracleDataAdapter

    Private dt1 As New DataTable    'sinvoice
    Private dt2 As New DataTable    'sinvoiced
    Private dt3 As New DataTable    'sinvoicev
    Private dt4 As New DataTable    'gaccdudate

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn

        Adaptadores()

    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal Numero As String)
        Me.New(cn)
        Abrir(Numero)
    End Sub
    Private Sub Adaptadores()
        Dim Sql As String

        Sql = "SELECT * FROM sinvoice WHERE num_0 = :num_0"
        da1 = New OracleDataAdapter(Sql, cn)
        da1.SelectCommand.Parameters.Add("num_0", OracleType.VarChar)
        da1.UpdateCommand = New OracleCommandBuilder(da1).GetUpdateCommand

        Sql = "SELECT * FROM sinvoiced WHERE num_0 = :num_0"
        da2 = New OracleDataAdapter(Sql, cn)
        da2.SelectCommand.Parameters.Add("num_0", OracleType.VarChar)

        Sql = "SELECT * FROM sinvoicev WHERE num_0 = :num_0"
        da3 = New OracleDataAdapter(Sql, cn)
        da3.SelectCommand.Parameters.Add("num_0", OracleType.VarChar)

        Sql = "SELECT * FROM gaccdudate WHERE num_0 = :num_0 AND dudlig_0 = 1"
        da4 = New OracleDataAdapter(Sql, cn)
        da4.SelectCommand.Parameters.Add("num_0", OracleType.VarChar)

    End Sub
    Public Function Abrir(ByVal Nro As String) As Boolean
        dt1.Clear()
        dt2.Clear()
        dt3.Clear()
        dt4.Clear()

        da1.SelectCommand.Parameters("num_0").Value = Nro
        da2.SelectCommand.Parameters("num_0").Value = Nro
        da3.SelectCommand.Parameters("num_0").Value = Nro

        da1.Fill(dt1)
        da2.Fill(dt2)
        da3.Fill(dt3)

        Return dt1.Rows.Count > 0

    End Function
    Public Function AbrirPorSolicitud(ByVal Nro As String) As Boolean
        'Busca la factura de una solicitud de servicio y la abre
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim flg As Boolean = False

        Sql = "SELECT * FROM sinvoicev WHERE sihori_0 = 8 AND sihorinum_0 = :sre"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("sre", OracleType.VarChar).Value = Nro
        da.Fill(dt)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            flg = Abrir(dr("num_0").ToString)
        End If

        da.Dispose()
        dt.Dispose()

        Return flg

    End Function
    Public Function ExisteArticulo(ByVal Codigo As String) As Boolean
        Dim flg As Boolean = False
        Dim dr As DataRow

        For Each dr In dt2.Rows
            If dr("itmref_0").ToString = Codigo Then
                flg = True
                Exit For
            End If
        Next

        Return flg

    End Function
    Public Function FacturaFiscalImpresa() As Boolean
        Dim dr As DataRow
        Try
            dr = dt1.Rows(0)
            dr.BeginEdit()
            dr("xfiscal_0") = 2
            dr.EndEdit()

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function
    Public Function EsFacturaElectronica() As Boolean
        Dim flg As Boolean = False
        Dim dr As DataRow

        If dt1.Rows.Count > 0 Then
            dr = dt1.Rows(0)

            flg = CInt(dr("xfacte_0")) = 3
        End If

        Return flg

    End Function
    Public Function Grabar() As Boolean
        Try
            da1.Update(dt1)
            Return True

        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function CheckCuit() As Boolean
        Select Case TipoDoc
            Case "80", "86"
                'Si el cuit tiene mas de 11 digitos envio error
                If Cuit.Length > 11 Then Return False

                Dim cuit2 As String = ""
                Dim n(9) As Integer
                Dim v1 As Integer = 0
                Dim v2 As Integer = 0
                Dim v3 As Integer = 0

                n(0) = 5 : n(1) = 4 : n(2) = 3 : n(3) = 2 : n(4) = 7
                n(5) = 6 : n(6) = 5 : n(7) = 4 : n(8) = 3 : n(9) = 2

                For i As Integer = 0 To Cuit.Length - 2
                    Dim c As String

                    c = Cuit.Substring(i, 1)

                    If IsNumeric(c) Then
                        Dim j As Integer = CInt(c)

                        cuit2 &= c
                        v1 += j * n(i)

                    Else
                        Exit For

                    End If

                Next

                v2 = v1 Mod 11
                v3 = 11 - v2

                Select Case v3
                    Case 11
                        cuit2 &= "0"
                    Case 10
                        cuit2 &= "9"
                    Case Else
                        cuit2 &= v3
                End Select

                Return Cuit = cuit2

            Case Else
                Return True

        End Select

    End Function
    Public Sub SetCAE(ByVal Cae As String, ByVal Vto As Date)
        Dim dr As DataRow = dt1.Rows(0)

        dr.BeginEdit()
        dr("xcaesta_0") = 2
        dr("xcae_0") = Cae
        dr("xdatvlycae_0") = Vto
        dr("xexpflg_0") = 2
        dr.EndEdit()

    End Sub
    Public ReadOnly Property CAE() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("xcae_0").ToString.Trim
            End If

            Return txt

        End Get
    End Property
    Public ReadOnly Property Sociedad() As Clases.Sociedad
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Dim s As New Sociedad(cn)
            s.abrir(dr("cpy_0").ToString)
            Return s
        End Get
    End Property
    Public ReadOnly Property CbteOrigenNumero() As String
        Get
            Dim dr As DataRow = dt3.Rows(0)

            Return dr("sihorinum_0").ToString

        End Get
    End Property
    Public ReadOnly Property CbteOrigen() As Factura
        Get
            Dim f As New Factura(cn)
            If f.Abrir(Me.CbteOrigenNumero) Then
                Return f
            Else
                Return Nothing
            End If
        End Get
    End Property
    Public ReadOnly Property Numero() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("num_0").ToString
        End Get
    End Property
    Public ReadOnly Property Cliente() As Cliente
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Dim bpc As New Cliente(cn)
            bpc.Abrir(dr("bpr_0").ToString)
            Return bpc
        End Get
    End Property
    Public ReadOnly Property UsuarioCreacion() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("creusr_0").ToString()
        End Get
    End Property
    Public ReadOnly Property CodigoVendedor() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("rep_0").ToString()
        End Get
    End Property
    Public ReadOnly Property ImporteII() As Double
        Get
            Dim dr As DataRow = dt1.Rows(0)

            Return CDbl(dr("amtati_0"))

        End Get
    End Property
    Public ReadOnly Property TipoIVA() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("vac_0").ToString

        End Get
    End Property
    Public ReadOnly Property DireccionFactura() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("bpaaddlig_0").ToString

        End Get
    End Property
    Public ReadOnly Property CiudadFactura() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("cty_0").ToString

        End Get
    End Property
    Public ReadOnly Property CondicionPagoFactura() As CondicionPago
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return New CondicionPago(cn, dr("pte_0").ToString)

        End Get
    End Property
    Public ReadOnly Property DetalleFactura() As DataTable
        Get
            Return dt2
        End Get
    End Property
    Public ReadOnly Property Planta() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("fcy_0").ToString

        End Get
    End Property
    Public ReadOnly Property Cuit() As String
        Get
            Dim txt As String = ""
            Dim dr As DataRow

            dr = dt1.Rows(0)
            txt = dr("docnum_0").ToString.Trim
            If txt.Length = 0 Then txt = "0"
            Return txt
        End Get
    End Property
    Public ReadOnly Property TipoDoc() As String
        Get
            Dim txt As String = ""
            Dim dr As DataRow

            dr = dt1.Rows(0)
            txt = dr("doctyp_0").ToString.Trim
            Return txt
        End Get
    End Property
    Public ReadOnly Property TipoComprobante() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("sivtyp_0").ToString
        End Get
    End Property
    Public ReadOnly Property LetraComprobante() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("clsvcr_0").ToString
        End Get
    End Property
    Public ReadOnly Property AFIPTipoDoc() As Integer
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return CInt(dr("xafipcod_0"))
        End Get
    End Property
    Public ReadOnly Property PuntoVenta() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("scuvcr_0").ToString
        End Get
    End Property
    Public ReadOnly Property TieneRemito() As String
        Get
            Dim dr As DataRow = dt2.Rows(0)
            Return dr("SIHORINUM_0").ToString
        End Get
    End Property
    Public ReadOnly Property Fecha() As Date
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return CDate(dr("accdat_0"))
        End Get
    End Property
    Public ReadOnly Property CantImpuestos() As Integer
        Get
            Dim dr As DataRow
            Dim i As Integer = 0

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                i = CInt(dr("nbrtax_0"))
            End If

            Return i
        End Get
    End Property
    Public ReadOnly Property TipoImpuesto(ByVal idx As Integer) As String
        Get
            Dim dr As DataRow
            Dim i As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                i = dr("tax_" & idx.ToString).ToString
            End If

            Return i
        End Get
    End Property
    Public ReadOnly Property ImpuestoImporte(ByVal idx As Integer) As Double
        Get
            Dim dr As DataRow
            Dim i As Double = 0

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                i = CDbl(dr("amttax_" & idx.ToString))
            End If

            Return i
        End Get
    End Property
    Public ReadOnly Property ImpuestoBase(ByVal idx As Integer) As Double
        Get
            Dim dr As DataRow
            Dim i As Double = 0

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                i = CDbl(dr("bastax_" & idx.ToString))
            End If

            Return i
        End Get
    End Property
    'Propiedades para Factura de Crédito Electrónica
    Public ReadOnly Property FechaVtoPago() As Date
        Get
            Dim dr As DataRow

            If dt4.Rows.Count = 0 Then
                da4.SelectCommand.Parameters("num_0").Value = Me.Numero
                da4.Fill(dt4)
                dr = dt4.Rows(0)
                Return CDate(dr("duddat_0"))
            End If

            Return Nothing

        End Get
    End Property
    Public ReadOnly Property BancoCodigo() As String
        Get
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                Return dr("xbanfce_0").ToString.Trim
            Else
                Return ""
            End If

        End Get
    End Property
    Public ReadOnly Property DivisaCodigo() As String
        Get
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                Return dr("cur_0").ToString.Trim
            Else
                Return ""
            End If

        End Get
    End Property
    Public ReadOnly Property Divisa() As Divisa
        Get
            Dim d As Divisa = Nothing

            If Me.DivisaCodigo <> "" Then
                d = New Divisa(cn)
                d.Abrir(Me.DivisaCodigo)
            End If

            Return d
        End Get
    End Property
    Public ReadOnly Property Banco() As Banco
        Get
            Dim ban As New Banco(cn)
            ban.Abrir(Me.BancoCodigo)
            Return ban
        End Get
    End Property
    Public ReadOnly Property Cotizacion() As Double
        Get
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                Return CDbl(dr("ratcur_0"))
            Else
                Return 1
            End If
        End Get
    End Property
    Public Property FCEAnula() As Boolean
        Get
            Dim b As Boolean = False

            If dt1.Rows.Count > 0 Then
                Dim dr As DataRow
                dr = dt1.Rows(0)
                b = CBool(IIf(CInt(dr("xanufce_0")) <> 2, False, True))
            End If

            Return b
        End Get
        Set(ByVal value As Boolean)
            If dt1.Rows.Count > 0 Then
                Dim dr As DataRow
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("xanufce_0") = IIf(value, 2, 1)
                dr.EndEdit()
            End If
        End Set
    End Property


End Class