Imports System.Data.OracleClient

Public Class Temporal
    Private dt As New DataTable
    Private da As OracleDataAdapter
    Private cn As OracleConnection
    Private dr As DataRow

    Private Tipo As String
    Private Usuario As String

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String

        Me.cn = cn
        Me.Tipo = Tipo
        Me.Usuario = Usuario

        Sql = "SELECT * FROM xtmp WHERE usr_0 = :usr AND tipo_0 = :tipo"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("usr", OracleType.VarChar)
        da.SelectCommand.Parameters.Add("tipo", OracleType.VarChar)

        Sql = "DELETE FROM xtmp WHERE usr_0 = :usr_0 AND tipo_0 = :tipo_0"
        da.DeleteCommand = New OracleCommand(Sql, cn)

        Sql = "INSERT INTO xtmp VALUES(:usr_0, :tipo_0,"
        Sql &= ":n_0,:n_1,:n_2,:n_3,:n_4,:n_5,:n_6,:n_7,:n_8,:n_9,:n_10,:n_11,:n_12,:n_13,:n_14,"
        Sql &= ":s_0,:s_1,:s_2,:s_3,:s_4,:s_5,:s_6,:s_7,:s_8,:s_9,:s_10,:s_11,:s_12,:s_13,:s_14,"
        Sql &= ":d_0,:d_1,:d_2,:d_3,:d_4,:d_5,:d_6,:d_7,:d_8,:d_9,:d_10,:d_11,:d_12,:d_13,:d_14)"
        da.InsertCommand = New OracleCommand(Sql, cn)

        Sql = "UPDATE xtmp SET n_0 = :n_0, n_1 = :n_1, n_2 = :n_2, n_3 = :n_3, n_4 = :n_4, n_5 = :n_5, n_6 = :n_6, n_7 = :n_7, n_8 = :n_8, n_9 = :n_9, n_10 = :n_10, n_11 = :n_11, n_12 = :n_12, n_13 = :n_13, n_14 = :n_14, "
        Sql &= "               s_0 = :s_0, s_1 = :s_1, s_2 = :s_2, s_3 = :s_3, s_4 = :s_4, s_5 = :s_5, s_6 = :s_6, s_7 = :s_7, s_8 = :s_8, s_9 = :s_9, s_10 = :s_10, s_11 = :s_11, s_12 = :s_12, s_13 = :s_13, s_14 = :s_14, "
        Sql &= "               d_0 = :d_0, d_1 = :d_1, d_2 = :d_2, d_3 = :d_3, d_4 = :d_4, d_5 = :d_5, d_6 = :d_6, d_7 = :d_7, d_8 = :d_8, d_9 = :d_9, d_10 = :d_10, d_11 = :d_11, d_12 = :d_12, d_13 = :d_13, d_14 = :d_14 "
        Sql &= "WHERE usr_0 = :usr_0 AND tipo_0 = :tipo_0"
        da.UpdateCommand = New OracleCommand(Sql, cn)

        Parametro(da.InsertCommand, "usr_0", OracleType.VarChar, DataRowVersion.Current)
        Parametro(da.InsertCommand, "tipo_0", OracleType.VarChar, DataRowVersion.Current)

        For i As Integer = 0 To 14
            Parametro(da.InsertCommand, "n_" & i.ToString, OracleType.Number, DataRowVersion.Current)
            Parametro(da.InsertCommand, "s_" & i.ToString, OracleType.VarChar, DataRowVersion.Current)
            Parametro(da.InsertCommand, "d_" & i.ToString, OracleType.DateTime, DataRowVersion.Current)

            Parametro(da.UpdateCommand, "n_" & i.ToString, OracleType.Number, DataRowVersion.Current)
            Parametro(da.UpdateCommand, "s_" & i.ToString, OracleType.VarChar, DataRowVersion.Current)
            Parametro(da.UpdateCommand, "d_" & i.ToString, OracleType.DateTime, DataRowVersion.Current)
        Next
        Parametro(da.UpdateCommand, "usr_0", OracleType.VarChar, DataRowVersion.Original)
        Parametro(da.UpdateCommand, "tipo_0", OracleType.VarChar, DataRowVersion.Original)

        Parametro(da.DeleteCommand, "usr_0", OracleType.VarChar, DataRowVersion.Original)
        Parametro(da.DeleteCommand, "tipo_0", OracleType.VarChar, DataRowVersion.Original)

        da.ContinueUpdateOnError = True
        da.FillSchema(dt, SchemaType.Mapped)

    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal Usr As Usuario, ByVal Tipo As String)
        Me.New(cn)
        Me.Usuario = Usr.Codigo
        Me.Tipo = Tipo
    End Sub

    Public Sub Nuevo()
        Nuevo(Me.Usuario, Me.Tipo)
    End Sub
    Public Sub Nuevo(ByVal Usuario As String, ByVal Tipo As String)
        dr = dt.NewRow

        For i As Integer = 0 To 14
            dr("n_" & i.ToString) = 0
            dr("s_" & i.ToString) = " "
            dr("d_" & i.ToString) = #12/31/1599#
        Next

        dr("usr_0") = Usuario
        dr("tipo_0") = Tipo

        dt.Rows.Add(dr)

    End Sub
    Public Sub Abrir()
        da.SelectCommand.Parameters("usr").Value = Usuario
        da.SelectCommand.Parameters("tipo").Value = Tipo
        dt.Clear()
        Try
            da.Fill(dt)
        Catch ex As Exception
        End Try
    End Sub
    Public Sub Registro(ByVal idx As Integer)
        dr = dt.Rows(idx)
    End Sub
    Public Sub Grabar()
        Try
            da.Update(dt)
        Catch ex As Exception

        End Try
    End Sub
    Public Sub LimpiarTabla()
        For Each dr As DataRow In dt.Rows
            If dr.RowState <> DataRowState.Deleted Then dr.Delete()
        Next
        'dt.Rows.Clear()

    End Sub
    Public Function Buscar(ByVal Idx As Integer, ByVal Numero As Double) As Boolean
        Dim dr As DataRow = Nothing
        Dim flg As Boolean = False

        For Each dr In dt.Rows
            If dr.RowState = DataRowState.Deleted Then Continue For
            If CDbl(dr("n_" & Idx.ToString)) = Numero Then
                Me.dr = dr
                flg = True
                Exit For
            End If
        Next

        Return flg
    End Function
    Public Property Numero(ByVal idx As Integer) As Double
        Get
            Return CDbl(dr("n_" & idx.ToString))
        End Get
        Set(ByVal value As Double)
            dr.BeginEdit()
            dr("n_" & idx.ToString) = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Cadena(ByVal idx As Integer) As String
        Get
            Return dr("s_" & idx.ToString).ToString
        End Get
        Set(ByVal value As String)
            dr.BeginEdit()
            dr("s_" & idx.ToString) = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Fecha(ByVal idx As Integer) As Date
        Get
            Return CDate(dr("d_" & idx.ToString))
        End Get
        Set(ByVal value As Date)
            dr.BeginEdit()
            dr("d_" & idx.ToString) = value
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property Count() As Integer
        Get
            Return dt.Rows.Count
        End Get
    End Property

End Class