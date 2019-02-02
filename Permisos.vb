Imports System.Data.OracleClient

Public Class Permisos
    Private cn As OracleConnection
    Private da1 As OracleDataAdapter
    Private da2 As OracleDataAdapter
    Private dt1 As New DataTable
    Private dt2 As New DataTable

    Public Sub New(ByVal cn As OracleConnection, ByVal Usr As String)
        Me.cn = cn

        Adaptadores()

        Abrir(Usr)

    End Sub
    Private Sub Adaptadores()
        Dim Sql As String

        Sql = "SELECT * FROM xnetper WHERE usr_0 = :usr"
        da1 = New OracleDataAdapter(Sql, cn)
        da1.SelectCommand.Parameters.Add("usr", OracleType.VarChar)

        Sql = "SELECT * FROM xnetvenc WHERE usr_0 = :usr"
        da2 = New OracleDataAdapter(Sql, cn)
        da2.SelectCommand.Parameters.Add("usr", OracleType.VarChar)

    End Sub
    Public Sub Abrir(ByVal usr As String)

        Try
            da1.SelectCommand.Parameters("usr").Value = usr
            da2.SelectCommand.Parameters("usr").Value = usr

            dt1.Clear()
            dt2.clear()

            da1.Fill(dt1)
            da2.fill(dt2)

        Catch ex As Exception

        End Try
    End Sub
    Private Function BuscarFuncion(ByVal Fnc As Integer) As Boolean
        Dim drx As DataRow = Nothing
        Dim flg As Boolean = False

        For Each dr As DataRow In dt1.Rows
            If CInt(dr("fnc_0")) = Fnc Then
                flg = True
                Exit For
            End If
        Next

        Return flg

    End Function
    Public ReadOnly Property AccesoFuncion(ByVal Fnc As Integer) As Boolean
        Get
            Return BuscarFuncion(Fnc)
        End Get

    End Property
    Public ReadOnly Property AccesoSecundario(ByVal Fnc As Integer, ByVal Cod As String) As Boolean
        Get
            Dim drx As DataRow = Nothing
            Dim flg As Boolean = False

            For Each dr As DataRow In dt1.Rows
                If CInt(dr("fncid_0")) = Fnc Then
                    Dim x As String = dr("flags_0").ToString
                    flg = dr("flags_0").ToString.Contains(Cod)
                End If
            Next

            Return flg
        End Get
    End Property
    Public Function TieneAccesoClienteAbonado(ByVal Rep As String) As Boolean
        Dim dr As DataRow
        Dim flg As Boolean = False

        For Each dr In dt2.Rows
            If dr("rep_0").ToString = Rep Then
                If CInt(dr("abo_0")) = 1 Then
                    flg = True
                End If
                Exit For
            End If
        Next

        Return flg

    End Function
    Public Function TieneAccesoClienteNoAbonado(ByVal Rep As String) As Boolean
        Dim dr As DataRow
        Dim flg As Boolean = False

        For Each dr In dt2.Rows
            If dr("rep_0").ToString = Rep Then
                If CInt(dr("noabo_0")) = 1 Then
                    flg = True
                End If
                Exit For
            End If
        Next

        Return flg
    End Function
End Class
