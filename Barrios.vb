Imports System.Data.OracleClient
Imports System.Windows.Forms

Public Class Barrios
    Private cn As OracleConnection
    Private dt1 As New DataTable 'zbarrios
    Private dt2 As New DataTable 'zbarrioscp
    Private da1 As OracleDataAdapter
    Private da2 As OracleDataAdapter
    Private Barrio As Integer = 0

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String

        Me.cn = cn

        Sql = "SELECT * FROM zbarrios where id_0 = :id"
        da1 = New OracleDataAdapter(Sql, cn)
        da1.SelectCommand.Parameters.Add("id", OracleType.Number)
        da1.InsertCommand = New OracleCommandBuilder(da1).GetInsertCommand
        da1.UpdateCommand = New OracleCommandBuilder(da1).GetUpdateCommand
        da1.DeleteCommand = New OracleCommandBuilder(da1).GetDeleteCommand

        Sql = "SELECT * FROM zbarrioscp where id_0 = :id"
        da2 = New OracleDataAdapter(Sql, cn)
        da2.SelectCommand.Parameters.Add("id", OracleType.Number)
        da2.InsertCommand = New OracleCommandBuilder(da2).GetInsertCommand
        da2.UpdateCommand = New OracleCommandBuilder(da2).GetUpdateCommand
        da2.DeleteCommand = New OracleCommandBuilder(da2).GetDeleteCommand

    End Sub
    Public Sub Abrir(ByVal Barrio As Integer)
        Me.Barrio = Barrio

        da1.SelectCommand.Parameters("id").Value = Barrio
        da2.SelectCommand.Parameters("id").Value = Barrio

        dt1.Clear()
        dt2.Clear()

        da1.Fill(dt1)
        da2.Fill(dt2)

    End Sub
    Public Sub AgregarCP(ByVal cp As String)
        Dim dr As DataRow

        For Each dr In dt2.Rows
            If dr("poscod_0").ToString.Trim = cp.Trim Then Exit Sub
        Next

        dr = dt2.NewRow
        dr("id_0") = Barrio
        dr("poscod_0") = cp
        dt2.Rows.Add(dr)
        da2.Update(dt2)

    End Sub
    Public Sub BarriosTodos(ByVal lst As ListBox)
        Dim Sql As String
        Dim dt As New DataTable
        Dim da As OracleDataAdapter

        Sql = "select * from zbarrios order by nombre_0"
        da = New OracleDataAdapter(Sql, cn)
        da.Fill(dt)
        da.Dispose()

        With lst
            .ValueMember = "id_0"
            .DisplayMember = "nombre_0"
            .DataSource = dt
        End With
    End Sub
    Public Sub BarriosPorCP(ByVal CodigoPostal As String, ByVal lst As ListBox)
        Dim Sql As String
        Dim dt As New DataTable
        Dim da As OracleDataAdapter

        Sql = "select zb.* "
        Sql &= "from zbarrios zb inner join zbarrioscp zbp on (zb.id_0 = zbp.id_0) "
        Sql &= "where poscod_0 = :cp "
        Sql &= "order by nombre_0"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("cp", OracleType.VarChar).Value = CodigoPostal
        da.Fill(dt)
        da.Dispose()

        With lst
            .ValueMember = "id_0"
            .DisplayMember = "nombre_0"
            .DataSource = dt
            '.SelectedItems.Clear()
        End With

    End Sub

    Public ReadOnly Property ZonaCamioneta() As Integer
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CInt(dr("camioneta_0"))
        End Get
    End Property
    Public ReadOnly Property ZonaCourier() As Integer
        Get
            Dim dr As DataRow
            dr = dt1.Rows(0)
            Return CInt(dr("courier_0"))
        End Get
    End Property
End Class