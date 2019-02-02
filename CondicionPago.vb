Imports System.Data.OracleClient
Imports System.Windows.Forms

Public Class CondicionPago
    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String

        Me.cn = cn

        Sql = "SELECT * FROM tabpayterm WHERE ptelin_0 = 1 AND pte_0 = :pte_0"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("pte_0", OracleType.VarChar)

    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal Codigo As String)
        Me.New(cn)
        EstablecerCodigo(Codigo)
    End Sub
    Public Sub EstablecerCodigo(ByVal Codigo As String)
        da.SelectCommand.Parameters("pte_0").Value = Codigo
        dt = New DataTable
        da.Fill(dt)
    End Sub
    Public ReadOnly Property Codigo() As String
        Get
            Dim dr As DataRow
            Dim pte As String = " "

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                pte = dr("pte_0").ToString
            End If

            Return pte
        End Get
    End Property
    Public ReadOnly Property Prioridad() As Integer
        Get
            Return CInt(dt.Rows(0).Item("yprioridad_0"))
        End Get
    End Property
    Public ReadOnly Property Descripcion() As String
        Get
            Dim txt As String = dt.Rows(0).Item("landessho_0").ToString
            Dim p1 As Integer = txt.IndexOf("~"c)
            Dim p2 As Integer = txt.IndexOf("~"c, p1 + 1)
            txt = txt.Substring(p1 + 1, p2 - p1 - 1)
            Return txt
        End Get
    End Property
    Shared Sub LlenarComboBox(ByVal cn As OracleConnection, ByVal cbo As ComboBox, Optional ByVal Blanco As Boolean = False)
        Dim spl() As String
        Dim sql As String = "SELECT pte_0, landessho_0, (pte_0 || ' - ' || landessho_0) as completo FROM tabpayterm WHERE ptelin_0 = 1 order by pte_0"
        Dim da As New OracleDataAdapter(sql, cn)
        Dim dt As New DataTable
        Dim dr As DataRow
        da.Fill(dt)

        For Each dr In dt.Rows
            spl = Split(dr("landessho_0").ToString, "~")

            dr.BeginEdit()
            dr("landessho_0") = dr("pte_0").ToString & " - " & spl(1)
            dr.EndEdit()
        Next

        If Blanco Then
            dr = dt.NewRow
            dr(0) = " "
            dr(1) = " "
            dt.Rows.InsertAt(dr, 0)
        End If

        With cbo
            .DataSource = dt
            .DisplayMember = "landessho_0"
            .ValueMember = "pte_0"
        End With
    End Sub
    Public ReadOnly Property TieneDatos() As Boolean
        Get
            If dt Is Nothing Then
                Return False
            ElseIf dt.Rows.Count = 0 Then
                Return False
            Else
                Return True
            End If
        End Get
    End Property

End Class