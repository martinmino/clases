Imports System.Data.OracleClient

Public Class Pacto
    Private cn As OracleConnection
    Private dah As OracleDataAdapter
    Private dth As DataTable
    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String
        Me.cn = cn

        'Adaptador de tabla cabecera
        Sql = " Select * from xpacto where nro_0 = :nro_0"
        dah = New OracleDataAdapter(Sql, cn)
        dah.SelectCommand.Parameters.Add("nro_0", OracleType.Number)
        dah.InsertCommand = New OracleCommandBuilder(dah).GetInsertCommand
        dah.UpdateCommand = New OracleCommandBuilder(dah).GetUpdateCommand
        dah.DeleteCommand = New OracleCommandBuilder(dah).GetDeleteCommand

        dth = New DataTable
        dah.FillSchema(dth, SchemaType.Mapped)
    End Sub
    Public Sub Nuevo()
        Dim dr As DataRow

        dth.Clear()

        dr = dth.NewRow
        dr("nro_0") = 0
        dr("bpcnam_0") = " "
        dr("direccion_0") = " "
        dr("rep_0") = " "
        dr("cot_0") = " "
        dr("cty_0") = " "
        dr("colega_0") = " "
        dr("estado_0") = " "
        dr("credat_0") = DateTime.Now
        dr("Mail_0") = 0
        dth.Rows.Add(dr)

    End Sub
    Public Function Abrir(ByVal Nro As Long) As Boolean
        dth.Clear()
        dah.SelectCommand.Parameters("nro_0").Value = Nro
        dah.Fill(dth)

        Return dth.Rows.Count > 0

    End Function
    Public Function LLenarDt() As DataTable
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim Sql As String
        Me.cn = cn
        'Sql = " select * from xpacto"
        Sql = "select nro_0 as Numero, bpcnam_0  as Razon_Social,Direccion_0 as Direccion,rep_0 as Vendedor, Cot_0 as Cotizar,cty_0 as Ciudad, Colega_0 as Colega,Estado_0 as Estado,Credat_0 as Fecha_Creacion from xpacto order by nro_0"
        da = New OracleDataAdapter(Sql, cn)
        da.Fill(dt)
        Return dt
    End Function
    Public Function LlenarDtPendiente() As DataTable
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim Sql As String
        Me.cn = cn
        Sql = "select nro_0 as Numero, bpcnam_0  as Razon_Social,Direccion_0 as Direccion,rep_0 as Vendedor, Cot_0 as Cotizar,cty_0 as Ciudad, Colega_0 as Colega,Estado_0 as Estado,Credat_0 as Fecha_Creacion from xpacto"
        Sql &= " where estado_0 not in ('Liberado','Precio Colega')  order by nro_0 "
        da = New OracleDataAdapter(Sql, cn)
        da.Fill(dt)
        Return dt
    End Function
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
    Public Property NombreCliente() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("bpcnam_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("bpcnam_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Direccion() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("direccion_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("direccion_0") = IIf(value = "", " ", value)
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
            dr("rep_0") = IIf(value = "", " ", value)
            dr.EndEdit()
        End Set
    End Property
    Public Property Cotizacion() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("cot_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("cot_0") = IIf(value = "", " ", value)
            dr.EndEdit()
        End Set
    End Property
    Public Property Ciudad() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("cty_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("cty_0") = IIf(value = "", " ", value)
            dr.EndEdit()
        End Set
    End Property
    Public Property Estado() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("estado_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("estado_0") = IIf(value = "", " ", value)
            dr.EndEdit()
        End Set
    End Property
    Public Property Mail() As Integer
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return CInt(dr("Mail_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("mail_0") = value
            dr.EndEdit()
        End Set
    End Property

    Public Property Colega() As String
        Get
            Dim dr As DataRow
            dr = dth.Rows(0)
            Return dr("Colega_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow
            dr = dth.Rows(0)
            dr.BeginEdit()
            dr("Colega_0") = IIf(value = "", " ", value)
            dr.EndEdit()
        End Set
    End Property
    Public Function grabar() As Boolean
        Dim n As Long = Me.Numero
        Dim i As Integer = 0
        Dim dr As DataRow

        dr = dth.Rows(0)

        dr.BeginEdit()
        If n = 0 Then
            n = ProximoNumero()
            dr("nro_0") = n
        End If
        dr.EndEdit()

        Try
            dah.Update(dth)

        Catch ex As Exception
            MsgBox(ex.ToString)
            Return False

        End Try

        Return True

    End Function
    Public Function GrabarModificacion() As Boolean
        Try
            dah.Update(dth)

        Catch ex As Exception
            MsgBox(ex.ToString)
            Return False

        End Try

        Return True
    End Function
    Private Function ProximoNumero() As Integer
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim Nro As Integer = 0

        da = New OracleDataAdapter("SELECT MAX(nro_0) FROM xpacto", cn)
        da.Fill(dt)
        da.Dispose()

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            If Not IsDBNull(dr(0)) Then Nro = CInt(dr(0))
        End If

        Nro += 1

        Return Nro

    End Function
End Class
