Imports System.Data.OracleClient

Public Class Licitacion
    Private cn As OracleConnection
    Private da1 As OracleDataAdapter
    Private dt1 As New DataTable
    Private da2 As OracleDataAdapter
    Private dt2 As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
        Adaptadores()

        dt1 = New DataTable
        da1.FillSchema(dt1, SchemaType.Mapped)

        dt2 = New DataTable
        da2.FillSchema(dt2, SchemaType.Mapped)

    End Sub
    Private Sub Adaptadores()
        Dim sql As String

        sql = "SELECT * FROM xlicita WHERE nro_0 = :nro"
        da1 = New OracleDataAdapter(sql, cn)
        da1.SelectCommand.Parameters.Add("nro", OracleType.VarChar)
        da1.InsertCommand = New OracleCommandBuilder(da1).GetInsertCommand
        da1.DeleteCommand = New OracleCommandBuilder(da1).GetDeleteCommand
        da1.UpdateCommand = New OracleCommandBuilder(da1).GetUpdateCommand

        sql = "SELECT * FROM xlicitad WHERE nro_0 = :nro"
        da2 = New OracleDataAdapter(sql, cn)
        da2.SelectCommand.Parameters.Add("nro", OracleType.VarChar)
        da2.InsertCommand = New OracleCommandBuilder(da2).GetInsertCommand
        da2.DeleteCommand = New OracleCommandBuilder(da2).GetDeleteCommand
        da2.UpdateCommand = New OracleCommandBuilder(da2).GetUpdateCommand

    End Sub
    Public Sub Nueva()
        Dim dr As DataRow

        dt1.Clear()
        dt2.Clear()

        dr = dt1.NewRow
        dr("nro_0") = 0
        dr("bpcnum_0") = " "
        dr("bpcnam_0") = " "
        dr("bpaadd_0") = " "
        dr("bpaaddlig_0") = " "
        dr("cty_0") = " "
        dr("sqhdat_0") = #12/31/1599#
        dr("sqhnum_0") = " "
        dr("sqhamt_0") = 0 '" "
        dr("rep_0") = " "
        dr("nuevo_0") = 1
        dr("service_0") = 1
        dr("agua_0") = 1
        dr("deteccion_0") = 1
        dr("licitatyp_0") = 1
        dr("licitanum_0") = " "
        dr("apertura_0") = #12/31/1599#
        dr("compras_0") = " "
        dr("estado_0") = 0
        dr("adjudidat_0") = #12/31/1599#
        dr("obs_0") = " "
        dr("poliza_0") = " "
        dr("polizadat_0") = #12/31/1599#
        dr("polizavto_0") = #12/31/1599#
        dr("sihnum_0") = " "
        dr("sdhnum_0") = " "
        dt1.Rows.Add(dr)

    End Sub
    Public Function Abrir(ByVal Nro As Integer) As Boolean
        dt1.Clear()
        dt2.Clear()

        da1.SelectCommand.Parameters("nro").Value = Nro
        dt1.Clear()
        da1.Fill(dt1)

        da2.SelectCommand.Parameters("nro").Value = Nro
        dt2.Clear()
        da2.Fill(dt2)

        Return dt1.Rows.Count > 0
    End Function
    Public Sub Grabar()
        Dim dr As DataRow

        If dt1.Rows.Count > 0 Then
            dr = dt1.Rows(0)

            If Numero = 0 Then
                Dim i As Long = ProximoNumero()

                dr.BeginEdit()
                dr("nro_0") = i
                dr.EndEdit()

                For Each dr2 As DataRow In dt2.Rows
                    dr2.BeginEdit()
                    dr2("nro_0") = i
                    dr2.EndEdit()
                Next

            End If
        End If

        da1.Update(dt1)
        da2.Update(dt2)

    End Sub
    Public Sub setCliente(ByVal bpc As Cliente)
        Dim dr As DataRow

        If dt1.Rows.Count > 0 Then
            dr = dt1.Rows(0)

            dr.BeginEdit()
            dr("bpcnum_0") = bpc.Codigo
            dr("bpcnam_0") = bpc.Nombre
            dr("rep_0") = bpc.Vendedor1Codigo
            dr.EndEdit()

        End If
    End Sub
    Public Sub setSucursal(ByVal bpa As Sucursal)
        Dim dr As DataRow

        If dt1.Rows.Count > 0 Then
            dr = dt1.Rows(0)

            dr.BeginEdit()
            dr("bpaadd_0") = bpa.Sucursal
            dr("bpaaddlig_0") = bpa.Direccion
            dr("cty_0") = bpa.Ciudad
            dr.EndEdit()

        End If
    End Sub
    Public Sub AgregarDetalle(ByVal dt As DataTable)
        Dim i As Integer

        i = dt2.Rows.Count

        For Each dr As DataRow In dt.Rows
            Dim dr2 As DataRow

            dr2 = dt2.NewRow
            dr2("nro_0") = Me.Numero
            dr2("lig_0") = 1000 * (i + 1)
            dr2("itmref_0") = dr("itmref_0")
            dr2("itmdes1_0") = dr("itmdes1_0").ToString
            dr2("ganada_0") = 0
            dr2("qty_0") = dr("qty_0")
            dr2("qty_1") = dr("qty_0")
            dr2("amt_0") = CDbl(dr("netpriati_0")) * CDbl(dr("qty_0"))
            dr2("amt_1") = CDbl(dr("netpriati_0")) * CDbl(dr("qty_0"))
            dr2("empresa_0") = " "
            dt2.Rows.Add(dr2)

            i += 1

        Next

    End Sub
    Private Function ProximoNumero() As Long
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim Nro As Long = 0

        da = New OracleDataAdapter("SELECT MAX(nro_0) FROM xlicita", cn)
        da.Fill(dt)
        da.Dispose()

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            If Not IsDBNull(dr(0)) Then Nro = CLng(dr(0))
        End If
        Nro += 1
        Return Nro

    End Function

    Public ReadOnly Property Numero() As Integer
        Get
            Dim dr As DataRow
            Dim i As Integer = 0

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                i = CInt(dr("nro_0"))
            End If

            Return i
        End Get
    End Property
    Public ReadOnly Property Cliente() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("bpcnum_0").ToString
            End If

            Return txt.Trim
        End Get
    End Property
    Public ReadOnly Property ClienteNombre() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("bpcnam_0").ToString
            End If

            Return txt.Trim
        End Get
    End Property
    Public ReadOnly Property Vendedor() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("rep_0").ToString
            End If

            Return txt.Trim
        End Get
    End Property
    Public Property Agua() As Boolean
        Get
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                Dim i As Integer = 0

                dr = dt1.Rows(0)

                Return CInt(dr("agua_0")) = 2

            End If

            Return False
        End Get
        Set(ByVal value As Boolean)
            If dt1.Rows.Count > 0 Then
                Dim dr As DataRow
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("agua_0") = IIf(value, 2, 1)
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property Deteccion() As Boolean
        Get
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                Dim i As Integer = 0

                dr = dt1.Rows(0)

                Return CInt(dr("deteccion_0")) = 2

            End If

            Return False
        End Get
        Set(ByVal value As Boolean)
            If dt1.Rows.Count > 0 Then
                Dim dr As DataRow
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("deteccion_0") = IIf(value, 2, 1)
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property Nuevo() As Boolean
        Get
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                Dim i As Integer = 0

                dr = dt1.Rows(0)

                Return CInt(dr("nuevo_0")) = 2

            End If

            Return False
        End Get
        Set(ByVal value As Boolean)
            If dt1.Rows.Count > 0 Then
                Dim dr As DataRow
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("nuevo_0") = IIf(value, 2, 1)
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property Service() As Boolean
        Get
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                Dim i As Integer = 0

                dr = dt1.Rows(0)

                Return CInt(dr("service_0")) = 2

            End If

            Return False
        End Get
        Set(ByVal value As Boolean)
            If dt1.Rows.Count > 0 Then
                Dim dr As DataRow
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("service_0") = IIf(value, 2, 1)
                dr.EndEdit()
            End If
        End Set
    End Property
    Public ReadOnly Property Sucursal() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("bpaadd_0").ToString
            End If

            Return txt.Trim
        End Get
    End Property
    Public ReadOnly Property Domicilio() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("bpaaddlig_0").ToString
            End If

            Return txt.Trim
        End Get
    End Property
    Public ReadOnly Property Localidad() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("cty_0").ToString
            End If

            Return txt.Trim
        End Get
    End Property
    Public Property TipoLicitacion() As Integer
        Get
            Dim dr As DataRow
            Dim i As Integer = 0

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)

                i = CInt(dr("licitatyp_0"))
            End If

            Return i
        End Get
        Set(ByVal value As Integer)
            If dt1.Rows.Count > 0 Then
                Dim dr As DataRow
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("licitatyp_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property NumeroLicitacion() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("licitanum_0").ToString
            End If

            Return txt.Trim
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr("licitanum_0") = IIf(value.Trim = "", " ", value)
            End If

        End Set
    End Property
    Public Property NumeroPresupuesto() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("sqhnum_0").ToString
            End If

            Return txt
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)

                dr.BeginEdit()
                dr("sqhnum_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property FechaPresupuesto() As Date
        Get
            Dim dr As DataRow
            Dim d As Date

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                d = CDate(dr("sqhdat_0"))
            End If

            Return d
        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)

                dr.BeginEdit()
                dr("sqhdat_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property TotalPresupuesto() As Double
        Get
            Dim dr As DataRow
            Dim i As Double = 0

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                i = CDbl(dr("sqhamt_0"))
            End If

            Return i
        End Get
        Set(ByVal value As Double)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)

                dr.BeginEdit()
                dr("sqhamt_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property Estado() As Integer
        Get
            Dim dr As DataRow
            Dim i As Integer = 0

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                i = CInt(dr("estado_0"))
            End If

            Return i
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("estado_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property

End Class