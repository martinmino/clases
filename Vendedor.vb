Imports System.Data.OracleClient
Imports System.Windows.Forms

Public Class Vendedor
    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
        Adaptadores()
    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal Codigo As String)
        Me.New(cn)
        Abrir(Codigo)
    End Sub

    Private Sub Adaptadores()
        Dim Sql As String

        Sql = "SELECT * FROM salesrep WHERE repnum_0 = :repnum_0"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("repnum_0", OracleType.VarChar)

    End Sub

    Public Function Abrir(ByVal Codigo As String) As Boolean
        dt.Clear()

        da.SelectCommand.Parameters("repnum_0").Value = Codigo
        da.Fill(dt)

        Return dt.Rows.Count > 0

    End Function
    Shared Sub Enlazar(ByVal cn As OracleConnection, ByVal cbo As combobox, ByVal Blanco As Boolean)
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim dr As DataRow

        Sql = "SELECT repnum_0, repnam_0 FROM salesrep ORDER BY repnam_0"
        da = New OracleDataAdapter(Sql, cn)
        da.Fill(dt)
        da.Dispose()

        If Blanco Then
            dr = dt.NewRow
            dr("repnum_0") = " "
            dr("repnam_0") = " "
            dt.Rows.InsertAt(dr, 0)
            dt.AcceptChanges()
        End If

        cbo.DisplayMember = "repnam_0"
        cbo.ValueMember = "repnum_0"
        cbo.DataSource = dt

    End Sub
    Shared Sub Enlazar_sin_cobranzas(ByVal cn As OracleConnection, ByVal cbo As ComboBox, ByVal Blanco As Boolean)
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim dr As DataRow


        Sql = "SELECT repnum_0, repnam_0, (repnum_0 || ' - ' || repnam_0) as completo FROM salesrep where (repnum_0 <> 'C001' and repnum_0 <> 'C002' and repnum_0 <> 'C003' and repnum_0 <> 'C004' and repnum_0 <> 'C005' and repnum_0 <> 'C006' and repnum_0 <> 'C007') ORDER BY repnum_0 "
        da = New OracleDataAdapter(Sql, cn)
        da.Fill(dt)
        da.Dispose()

        If Blanco Then
            dr = dt.NewRow
            dr("repnum_0") = " "
            dr("repnam_0") = " "
            dt.Rows.InsertAt(dr, 0)
            dt.AcceptChanges()
        End If


        cbo.DisplayMember = "completo"
        'cbo.DisplayMember = "repnam_0"
        cbo.ValueMember = "repnum_0"
        cbo.DataSource = dt

    End Sub
 
    Public ReadOnly Property Codigo() As String
        Get
            Dim txt As String = " "
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                txt = dr("repnum_0").ToString
            End If

            Return txt
        End Get
    End Property
    Public ReadOnly Property Nombre() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            If dt.Rows.Count = 0 Then Return " "
            Return dr("repnam_0").ToString
        End Get
    End Property
    Public ReadOnly Property Mail() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("xmail_0").ToString
        End Get
    End Property
    Public ReadOnly Property Analista() As Usuario
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return New Usuario(cn, dr("xanalis_0").ToString)
        End Get
    End Property
    Public ReadOnly Property Conexion() As OracleConnection
        Get
            Return cn
        End Get
    End Property
    Public ReadOnly Property Interno() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("xinterno_0").ToString.Trim
        End Get
    End Property
    Public ReadOnly Property Telefono() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("tel_0").ToString.Trim
        End Get
    End Property

End Class