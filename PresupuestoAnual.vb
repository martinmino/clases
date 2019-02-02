Imports System.Data.OracleClient

Public Class PresupuestoAnual
    Private da As OracleDataAdapter
    Private dt As New DataTable
    Private dt2 As New DataTable
    Private dr As DataRow
    Private cn As OracleConnection

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn

    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal Codigo As String, ByVal cuenta As Integer, ByVal ano As Integer)
        Me.New(cn)

        Abrir(Codigo, cuenta, ano)
    End Sub
    Public Function Abrir(ByVal Ano As Integer) As Boolean
        Dim Sql As String

        Sql = "SELECT * "
        Sql &= "FROM presupuest "
        Sql &= "WHERE extract(year  from dat_0) = :ano"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("ano", OracleType.Number).Value = Ano
        dt.Clear()
        da.Fill(dt)
        da.Dispose()

        Return dt.Rows.Count > 0

    End Function
    Public Overridable Function Abrir(ByVal CentroCostos As String, ByVal cuenta As Integer, ByVal ano As Integer) As Boolean
        Dim Sql As String = "SELECT * FROM presupuest WHERE cce_0 = :cce_0 and acc_0 = :acc_0 and  extract(year  from dat_0) = :ano"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("cce_0", OracleType.VarChar).Value = CentroCostos
        da.SelectCommand.Parameters.Add("acc_0", OracleType.Number).Value = cuenta
        da.SelectCommand.Parameters.Add("ano", OracleType.Number).Value = ano
        dt.Clear()
        da.Fill(dt)

        If dt.Rows.Count = 1 Then dr = dt.Rows(0)

        Return dt.Rows.Count <> 0

    End Function
    Public Function ObtenerValorAcumulado(ByVal CentroCosto As String, ByVal Mes As Integer) As Double
        Dim dv As New DataView(dt)
        Dim dr As DataRow
        Dim Suma As Double = 0

        'Filtro por Centro de Costo y Nro de Cuenta
        dv.RowFilter = "cce_0 = '" & CentroCosto & "'"

        For i = 0 To dv.Count - 1
            dr = dv.Item(i).Row

            For j = 0 To Mes - 1
                Suma += CDbl(dr("mes_" & j.ToString))
            Next

        Next

        Return Suma

    End Function
    Public Function ObtenerValorAcumulado(ByVal CentroCosto As String, ByVal Cuenta As String, ByVal Mes As Integer) As Double
        Dim dv As New DataView(dt)
        Dim dr As DataRow
        Dim Suma As Double = 0

        'Filtro por Centro de Costo y Nro de Cuenta
        dv.RowFilter = "cce_0 = '" & CentroCosto & "' AND acc_0 = '" & Cuenta & "'"

        If dv.Count > 0 Then
            dr = dv.Item(0).Row

            For i = 0 To Mes - 1
                Suma += CDbl(dr("mes_" & i.ToString))
            Next

        End If

        Return Suma

    End Function
    Public Overridable Function Abrir(ByVal CentroCostos As String, ByVal ano As Integer) As Boolean
        Dim Sql As String
        dt.Clear()
        Sql = "SELECT sum(mes_0) as mes_0,sum(mes_1)as mes_1,sum(mes_2)as mes_2,sum(mes_3)as mes_3,sum(mes_4)as mes_4,sum(mes_5)as mes_5,sum(mes_6)as mes_6,sum(mes_7)as mes_7,sum(mes_8)as mes_8,sum(mes_9)as mes_9,sum(mes_10)as mes_10,sum(mes_11)as mes_11 "
        Sql &= " FROM presupuest WHERE cce_0 = :cce_0 and  extract(year  from dat_0) = :ano AND ACC_0 <> '411106'"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("cce_0", OracleType.VarChar).Value = CentroCostos
        da.SelectCommand.Parameters.Add("ano", OracleType.Number).Value = ano

        da.Fill(dt)

        If dt.Rows.Count = 1 Then dr = dt.Rows(0)

        Return dt.Rows.Count <> 0

    End Function
    Public ReadOnly Property valor_mes(ByVal idx As Integer) As Double
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CDbl(dr("mes_" & idx.ToString))
        End Get
    End Property

    Public ReadOnly Property valor_mes_suma(ByVal idx As Integer) As Double
        Get
            Dim dr As DataRow = dt.Rows(0)
            ' MsgBox(CDbl(dr("mes_" & idx.ToString)))
            Return CDbl(dr("mes_" & idx.ToString))
        End Get
    End Property

End Class
