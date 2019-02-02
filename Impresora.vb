Imports System.Data.OracleClient

Public Class Impresora
    Private cn As OracleConnection
    Private da1 As OracleDataAdapter
    Private da2 As OracleDataAdapter
    Private dt1 As New DataTable
    Private dt2 As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn

        Dim Sql As String
        Sql = "SELECT * FROM aprinter WHERE cod_0 = :cod"
        da1 = New OracleDataAdapter(Sql, cn)
        da1.SelectCommand.Parameters.Add("cod", OracleType.VarChar)
        da1.UpdateCommand = New OracleCommandBuilder(da1).GetUpdateCommand

        Sql = "SELECT * FROM aprinterd WHERE cod_0 = :cod"
        da2 = New OracleDataAdapter(Sql, cn)
        da2.SelectCommand.Parameters.Add("cod", OracleType.VarChar)
        da2.UpdateCommand = New OracleCommandBuilder(da2).GetUpdateCommand

    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal prn As String)
        Me.New(cn)
        Abrir(prn)
    End Sub
    Public Function Abrir(ByVal prn As String) As Boolean
        da1.SelectCommand.Parameters("cod").Value = prn
        da2.SelectCommand.Parameters("cod").Value = prn
        dt1.Clear()
        dt2.Clear()

        Try
            da1.Fill(dt1)
            da2.Fill(dt2)

        Catch ex As Exception

        End Try

        Return dt1.Rows.Count > 0

    End Function
    Public Sub Grabar()
        Try
            da1.Update(dt1)
            da2.Update(dt2)

        Catch ex As Exception

        End Try
    End Sub
    Public Function GetDestinos() As DataTable
        'Devuelve todos los destinos
        Dim Sql As String = "SELECT * FROM aprinter ORDER BY cod_0"
        Dim da As New OracleDataAdapter(Sql, cn)
        Dim dt As New DataTable

        da.Fill(dt)
        da.Dispose()
        Return dt

    End Function
    Public ReadOnly Property RecursoRed() As String
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return dr("xprn_0").ToString
        End Get
    End Property
    Public Property X(ByVal campo As String) As Integer
        Get
            Dim v As Integer = 0

            For Each dr As DataRow In dt2.Rows
                If dr("campo_0").ToString = campo Then v = CInt(dr("x_0"))
            Next
            Return v

        End Get
        Set(ByVal value As Integer)
            For Each dr As DataRow In dt2.Rows
                If dr("campo_0").ToString = campo Then
                    dr.BeginEdit()
                    dr("x_0") = value
                    dr.EndEdit()
                End If
            Next
        End Set
    End Property
    Public Property Y(ByVal campo As String) As Integer
        Get
            Dim v As Integer = 0

            For Each dr As DataRow In dt2.Rows
                If dr("campo_0").ToString = campo Then v = CInt(dr("y_0"))
            Next

            Return v
        End Get
        Set(ByVal value As Integer)
            For Each dr As DataRow In dt2.Rows
                If dr("campo_0").ToString = campo Then
                    dr.BeginEdit()
                    dr("y_0") = value
                    dr.EndEdit()
                End If
            Next
        End Set
    End Property
    Public ReadOnly Property XY(ByVal campo As String) As String
        Get
            Return X(campo).ToString & "," & Y(campo).ToString
        End Get
    End Property
    Public ReadOnly Property Nombre() As String
        Get
            Dim dr = dt1.Rows(0)
            Return dr("des_0").ToString
        End Get
    End Property

End Class
