Imports System.Data.OracleClient

Public Class Numerador

    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable
    Private code As String = " "
    Private site As String = " "
    Private periode As Integer = 0

    Public Sub New(ByVal cn As OracleConnection, ByVal Codigo As String)
        Me.cn = cn

        Dim Sql As String
        Sql = "SELECT * FROM avalnum WHERE codnum_0 = :codnum"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("codnum", OracleType.VarChar).Value = Codigo
        Adaptador()

        code = Codigo

    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal Codigo As String, ByVal Sociedad As String)
        Me.cn = cn

        Dim Sql As String
        Sql = "SELECT * FROM avalnum WHERE codnum_0 = :codnum AND site_0 = :site"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("codnum", OracleType.VarChar).Value = Codigo
        da.SelectCommand.Parameters.Add("site", OracleType.VarChar).Value = Sociedad
        Adaptador()

        code = Codigo
        site = Sociedad

    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal Codigo As String, ByVal Sociedad As String, ByVal Periodo As Integer)
        Me.cn = cn

        Dim Sql As String
        Sql = "SELECT * FROM avalnum WHERE codnum_0 = :codnum AND site_0 = :site AND periode_0 = :periode"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("codnum", OracleType.VarChar).Value = Codigo
        da.SelectCommand.Parameters.Add("site", OracleType.VarChar).Value = Sociedad
        da.SelectCommand.Parameters.Add("periode", OracleType.VarChar).Value = Periodo
        Adaptador()

        code = Codigo
        site = Sociedad
        periode = Periodo

    End Sub
    Private Sub Adaptador()
        da.InsertCommand = New OracleCommandBuilder(da).GetInsertCommand
        da.DeleteCommand = New OracleCommandBuilder(da).GetDeleteCommand
        da.UpdateCommand = New OracleCommandBuilder(da).GetUpdateCommand
    End Sub
    Private Sub Abrir()
        Dim dr As DataRow

        da.Fill(dt)

        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
            dr("codnum_0") = code
            dr("site_0") = site
            dr("periode_0") = periode
            dr("comp_0") = " "
            dr("valeur_0") = 1
            dt.Rows.Add(dr)
        End If

    End Sub
    Public Function Valor() As Long
        'Devuelve el valor y lo incrementa en 1
        Dim dr As DataRow
        Dim v As Long = 0

        Try
            Abrir()

            dr = dt.Rows(0)
            v = CLng(dr("valeur_0"))
            dr.BeginEdit()
            dr("valeur_0") = v + 1
            dr.EndEdit()

            da.Update(dt)

        Catch ex As Exception
            v = -1

        End Try

        Return v

    End Function
    Public ReadOnly Property Periodo() As Integer
        Get
            Return periode
        End Get
    End Property

End Class