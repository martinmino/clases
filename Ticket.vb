Imports System.Data.OracleClient

Public Class Ticket

    Private cn As OracleConnection
    Private da1 As OracleDataAdapter 'XTICKETH
    Private da2 As OracleDataAdapter 'XTICKETD
    Private dt1 As New DataTable
    Public dt2 As New DataTable

    Private bpc As Cliente = Nothing
    Private bpa As Sucursal = Nothing

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
        Adaptadores()
    End Sub

    Private Sub Adaptadores()
        Dim Sql As String

        Sql = "SELECT * FROM xticketh WHERE nro_0 = :nro"
        da1 = New OracleDataAdapter(Sql, cn)
        da1.SelectCommand.Parameters.Add("nro", OracleType.Number)
        da1.InsertCommand = New OracleCommandBuilder(da1).GetInsertCommand
        da1.UpdateCommand = New OracleCommandBuilder(da1).GetUpdateCommand
        da1.DeleteCommand = New OracleCommandBuilder(da1).GetDeleteCommand

        Sql = "SELECT * FROM xticketd WHERE nro_0 = :nro"
        da2 = New OracleDataAdapter(Sql, cn)
        da2.SelectCommand.Parameters.Add("nro", OracleType.Number)
        da2.InsertCommand = New OracleCommandBuilder(da2).GetInsertCommand
        da2.UpdateCommand = New OracleCommandBuilder(da2).GetUpdateCommand
        da2.DeleteCommand = New OracleCommandBuilder(da2).GetDeleteCommand

        Try
            da1.FillSchema(dt1, SchemaType.Mapped)
            da2.FillSchema(dt2, SchemaType.Mapped)

        Catch ex As Exception

        End Try
    End Sub
    Public Sub AgregarComentario(ByVal txt As String, ByVal Usuario As String)
        Dim i As Integer = 0
        Dim dr As DataRow

        'Busco el siguiente numero de linea
        For Each dr In dt2.Rows
            If CInt(dr("lig_0")) > i Then
                i = CInt(dr("lig_0"))
            End If
        Next
        i += 1

        dr = dt2.NewRow
        dr("nro_0") = Me.Numero
        dr("lig_0") = i
        dr("dat_0") = Today
        dr("hora_0") = Now.ToString("HHmm")
        dr("usr_0") = Usuario
        dr("msg_0") = txt
        dt2.Rows.Add(dr)

        'Si es el primer comentario, cargo los campos de la cabecera
        If dt2.Rows.Count = 1 Then
            Me.Comentario = txt
        End If

    End Sub
    Public Sub Nuevo()
        Dim dr As DataRow

        bpc = Nothing
        bpa = Nothing

        dt1.Clear()
        dt2.Clear()

        dr = dt1.NewRow
        dr("nro_0") = 0
        dr("creusr_0") = " "
        dr("bpcnum_0") = " "
        dr("bpaadd_0") = " "
        dr("motivo_0") = 0
        dr("credat_0") = Date.Today
        dr("hora_0") = Now.ToString("HHmm")
        dr("asigusr_0") = " "
        dr("asigdat_0") = #12/31/1599#
        dr("estado_0") = 1
        dr("sdhnum_0") = " "
        dr("sdhnum_1") = " "
        dr("sdhnum_2") = " "
        dr("sihnum_0") = " "
        dr("sihnum_1") = " "
        dr("sihnum_2") = " "
        dr("itnnum_0") = " "
        dr("itnnum_1") = " "
        dr("itnnum_2") = " "
        dr("sohnum_0") = " "
        dr("sohnum_1") = " "
        dr("sohnum_2") = " "
        dr("contacto_0") = " "
        dr("tel_0") = " "
        dr("web_0") = " "
        dr("msg_0") = " "

        dt1.Rows.Add(dr)

    End Sub
    Private Function SiguienteNumero() As Integer
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim i As Integer = 0

        Sql = "SELECT MAX(nro_0) as nro FROM xticketh"
        da = New OracleDataAdapter(Sql, cn)
        da.Fill(dt)
        da.Dispose()

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            If Not IsDBNull(dr(0)) Then i = CInt(dr(0))
        End If

        i += 1

        Return i

    End Function
    Public Function Abrir(ByVal Nro As Integer) As Boolean
        Try
            da1.SelectCommand.Parameters("nro").Value = Nro
            da2.SelectCommand.Parameters("nro").Value = Nro

            dt1.Clear()
            dt2.Clear()

            da1.Fill(dt1)
            da2.Fill(dt2)

        Catch ex As Exception

        End Try

        Return dt1.Rows.Count > 0

    End Function
    Public Function Grabar() As Boolean
        Try
            If Me.Numero = 0 Then
                Me.Numero = SiguienteNumero()
                'Me.Fecha = Today
                'Me.Hora = Now.ToString("HHmm")

                For Each dr As DataRow In dt2.Rows
                    dr.BeginEdit()
                    dr("nro_0") = Me.Numero
                    dr.EndEdit()
                Next
            End If

            da1.Update(dt1)
            da2.Update(dt2)

            Return True

        Catch ex As Exception
            Return False

        End Try

    End Function
    Public Function AsignarA(ByVal Sector As String) As Boolean
        Me.AsignadoA = Sector
        Me.FechaAsignacion = Date.Today
    End Function
    Public Property Numero() As Integer
        Get
            Dim dr As DataRow
            Dim i As Integer = 0

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                i = CInt(dr("nro_0"))
            End If

            Return i
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("nro_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property Tipo() As Integer
        Get
            Dim dr As DataRow
            Dim i As Integer = 0

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                i = CInt(dr("motivo_0"))
            End If

            Return i
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("motivo_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public ReadOnly Property Cliente() As Cliente
        Get
            If bpc Is Nothing Then
                bpc = New Cliente(cn)
                bpc.Abrir(Me.ClienteCodigo)
            End If

            Return bpc
        End Get
    End Property
    Public Property ClienteCodigo() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("bpcnum_0").ToString
            End If

            Return txt
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("bpcnum_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public ReadOnly Property Sucursal() As Sucursal
        Get
            If bpa Is Nothing Then
                bpa = New Sucursal(cn)
                bpa.Abrir(Me.ClienteCodigo, Me.SucursalCodigo)
            End If

            Return bpa
        End Get
    End Property
    Public Property SucursalCodigo() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("bpaadd_0").ToString
            End If

            Return txt
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("bpaadd_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property Fecha() As Date
        Get
            Dim dr As DataRow
            Dim i As Date

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                i = CDate(dr("credat_0"))
            End If

            Return i
        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("credat_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property Hora() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("hora_0").ToString
            End If

            Return txt
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("hora_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property UsuarioCreacion() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("creusr_0").ToString
            End If

            Return txt
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("creusr_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property AsignadoA() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("asigusr_0").ToString
            End If

            Return txt
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("asigusr_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property FechaAsignacion() As Date
        Get
            Dim dr As DataRow
            Dim i As Date

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                i = CDate(dr("asigdat_0"))
            End If

            Return i
        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("asigdat_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property Comentario() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("msg_0").ToString
            End If

            Return txt
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("msg_0") = value
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
    Public Property Pedido(ByVal idx As Integer) As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("sohnum_" & idx.ToString).ToString
            End If

            Return txt
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("sohnum_" & idx.ToString) = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property Entrega(ByVal idx As Integer) As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("sdhnum_" & idx.ToString).ToString
            End If

            Return txt
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("sdhnum_" & idx.ToString) = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property Factura(ByVal idx As Integer) As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("sihnum_" & idx.ToString).ToString
            End If

            Return txt
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("sihnum_" & idx.ToString) = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property Intervencion(ByVal idx As Integer) As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("itnnum_" & idx.ToString).ToString
            End If

            Return txt
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("itnnum_" & idx.ToString) = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property Contacto() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("contacto_0").ToString
            End If

            Return txt
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("contacto_0") = IIf(value.Trim = "", " ", value.Trim)
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property Telefono() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("tel_0").ToString
            End If

            Return txt
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("tel_0") = IIf(value.Trim = "", " ", value.Trim)
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property Mail() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("web_0").ToString
            End If

            Return txt
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("web_0") = IIf(value.Trim = "", " ", value.Trim)
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property Observaciones() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("msg_0").ToString.Trim
            End If

            Return txt
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("msg_0") = IIf(value.Trim = "", " ", value.Trim)
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property Asignado() As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                txt = dr("asigusr_0").ToString.Trim
            End If

            Return txt
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("asigusr_0") = IIf(value.Trim = "", " ", value.Trim)
                dr.EndEdit()
            End If
        End Set
    End Property

End Class