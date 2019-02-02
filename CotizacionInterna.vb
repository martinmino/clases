Imports System.Data.OracleClient
Imports System.Windows.Forms
Imports System.IO
Imports System.Collections

Public Class CotizacionInterna

    Private cn As OracleConnection
    Private dah As OracleDataAdapter 'Adaptador Tabla Cabecera
    Private dad As OracleDataAdapter 'Adaptador Tabla Detalle
    Private dth As DataTable 'Tabla cabecera
    Public WithEvents dtd As DataTable 'Tabla detalle
    Private dtp As DataTable 'Tabla precios
    Private bpc As Cliente
    Private bpa As Sucursal

    Public Event PresupuestoModificado(ByVal sender As Object)
    'CONSTRUCTORES
    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String

        Me.cn = cn

        'Adaptador de tabla cabecera
        Sql = "SELECT * FROM xcotint WHERE nro_0 = :nro_0"
        dah = New OracleDataAdapter(Sql, cn)
        dah.SelectCommand.Parameters.Add("nro_0", OracleType.Number)
        dah.InsertCommand = New OracleCommandBuilder(dah).GetInsertCommand
        dah.UpdateCommand = New OracleCommandBuilder(dah).GetUpdateCommand
        dah.DeleteCommand = New OracleCommandBuilder(dah).GetDeleteCommand

        'Adaptador de tabla detalle
        Sql = "SELECT * FROM xcotintd WHERE nro_0 = :nro_0"
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
    Public Sub New(ByVal cn As OracleConnection, ByVal bpc As Cliente, ByVal cpy As String, ByVal EsH As Boolean, ByVal tipo As Integer)
        Me.cn = cn
        Nuevo()
    End Sub

    'SUB-PROGRAMAS
    Public Sub Nuevo()
        Dim dr As DataRow

        dtd.Clear()
        dth.Clear()

        dr = dth.NewRow
        dr("nro_0") = 0
        dr("dat_0") = Date.Today
        dr("dat_1") = Date.Today '#12/31/1599#
        dr("dat_2") = Date.Today '#12/31/1599#
        dr("bpcnum_0") = " "
        dr("estado_0") = 0
        dr("cotizador_0") = " "
        dr("obs_0") = " "
        dr("creusr_0") = USER

        dth.Rows.Add(dr)

        bpc = Nothing
        bpa = Nothing

    End Sub
    Public Function Abrir(ByVal Numero As String) As Boolean
        Dim flg As Boolean = False
        dth.Clear()
        dtd.Clear()
        dah.SelectCommand.Parameters("nro_0").Value = Numero
        dad.SelectCommand.Parameters("nro_0").Value = Numero

        Try
            dah.Fill(dth)
            dad.Fill(dtd)

            flg = dth.Rows.Count > 0

        Catch ex As Exception
            flg = False

        End Try

        Return flg

    End Function
    Private Function ProximoNumero() As Long
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim Nro As Long = 0

        da = New OracleDataAdapter("SELECT MAX(nro_0) FROM xcotint", cn)
        da.Fill(dt)
        da.Dispose()

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            If Not IsDBNull(dr(0)) Then Nro = CLng(dr(0))
        End If
        Nro += 1
        Return Nro

    End Function
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
    Public Sub EnlazarGrilla(ByVal dgv As DataGridView)
        dgv.DataSource = dtd
    End Sub


    Public Property Cliente() As String
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return dr("bpcnum_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dth.Rows(0)
            dr.BeginEdit()
            dr("bpcnum_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property cotizador() As String
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return dr("cotizador_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dth.Rows(0)
            dr.BeginEdit()
            dr("cotizador_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property observaciones() As String
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return dr("obs_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dth.Rows(0)
            dr.BeginEdit()
            dr("obs_0") = IIf(value = "", " ", value)
            dr.EndEdit()
        End Set
    End Property

    Public Property FechaCotizacion() As Date
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return CDate(dr("dat_0"))
        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow = dth.Rows(0)
            dr.BeginEdit()
            dr("dat_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property FechaRespuestaCotizacion() As Date
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return CDate(dr("dat_1"))
        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow = dth.Rows(0)
            dr.BeginEdit()
            dr("dat_1") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property FechaVigenciaPresupuesto() As Date
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return CDate(dr("dat_2"))
        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow = dth.Rows(0)
            dr.BeginEdit()
            dr("dat_2") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Estado() As Integer
        Get
            Dim dr As DataRow = dth.Rows(0)
            Return CInt(dr("estado_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow = dth.Rows(0)
            dr.BeginEdit()
            dr("estado_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Function Grabar() As Boolean
        Dim n As Long = Numero
        Dim i As Integer = 0
        Dim dr As DataRow

        'Recupero el numero del pedido actual
        dr = dth.Rows(0)
        dr.BeginEdit()
        If n = 0 Then
            n = ProximoNumero()
            dr("nro_0") = n
        End If
        'dr("estado_0") = 1

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
        'MsgBox("Se creo la cotizacion: " & n)
        dtd.Clear()
        dth.Clear()
        Return True

    End Function
    Public ReadOnly Property PrecioUnitario() As Double
        Get
            Dim dr As DataRow
            Dim p As Double = 0

            For Each dr In dtd.Rows
                If dr.RowState = DataRowState.Deleted Then Continue For
                p += CDbl(dr("precio_0"))
            Next

            Return p

        End Get
    End Property
    Public ReadOnly Property PrecioSugerido() As Double
        Get
            Dim dr As DataRow
            Dim p As Double = 0

            For Each dr In dtd.Rows
                If dr.RowState = DataRowState.Deleted Then Continue For
                p += CDbl(dr("precio_1"))
            Next

            Return p

        End Get
    End Property
    Private Sub dtd_RowChanged(ByVal sender As Object, ByVal e As System.Data.DataRowChangeEventArgs) Handles dtd.RowChanged
        RaiseEvent PresupuestoModificado(Me)
    End Sub
    Private Sub dtd_RowDeleted(ByVal sender As Object, ByVal e As System.Data.DataRowChangeEventArgs) Handles dtd.RowDeleted
        RaiseEvent PresupuestoModificado(Me)
    End Sub
    Private Sub dtd_TableNewRow(ByVal sender As Object, ByVal e As System.Data.DataTableNewRowEventArgs) Handles dtd.TableNewRow
        RaiseEvent PresupuestoModificado(Me)
    End Sub
End Class
