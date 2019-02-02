Imports System.Data.OracleClient

Public Class Feriados
    Private da As OracleDataAdapter
    Private cn As OracleConnection
    Private dt As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String

        Me.cn = cn

        Sql = "SELECT * FROM xferiados"
        da = New OracleDataAdapter(Sql, cn)

        da.InsertCommand = New OracleCommandBuilder(da).GetInsertCommand
        da.UpdateCommand = New OracleCommandBuilder(da).GetUpdateCommand
        da.DeleteCommand = New OracleCommandBuilder(da).GetDeleteCommand

        da.Fill(dt)

    End Sub
    Public Sub Add(ByVal Fecha As Date)
        Dim dr As DataRow

        dr = dt.NewRow
        dr("dat_0") = Fecha
        dt.Rows.Add(dr)
        da.Update(dt)

    End Sub
    Public Sub Delete(ByVal Fecha As Date)
        Dim dr As DataRow

        For Each dr In dt.Rows
            If Fecha = CDate(dr("dat_0")) Then
                dr.Delete()
                da.Update(dt)
                Exit For
            End If
        Next

    End Sub
    Public Function Existe(ByVal Fecha As Date) As Boolean
        Dim dr As DataRow
        Dim flg As Boolean = False

        For Each dr In dt.Rows
            If fecha = CDate(dr("dat_0")) Then
                flg = True
                Exit For
            End If
        Next

        Return flg

    End Function
    Public Function DiasHabilesMes(ByVal Fecha As Date) As Integer
        Dim Dias As Integer = 0
        Dim f As Date = New Date(Fecha.Year, Fecha.Month, 1)

        For i = 1 To Date.DaysInMonth(Fecha.Year, Fecha.Month)

            If f.DayOfWeek >= DayOfWeek.Monday And f.DayOfWeek <= DayOfWeek.Friday Then

                'Miro si no es feriado
                If Not Existe(f) Then Dias += 1

            End If

            f = f.AddDays(1)
        Next

        Return Dias

    End Function
    Public Function DiasHabilesHasta(ByVal Fecha As Date) As Integer

        Dim Dias As Integer = 0
        Dim f As Date = New Date(Fecha.Year, Fecha.Month, 1)

        For i = 1 To Date.DaysInMonth(Fecha.Year, Fecha.Month)

            If f.DayOfWeek >= DayOfWeek.Monday And f.DayOfWeek <= DayOfWeek.Friday Then

                'Miro si no es feriado
                If Not Existe(f) Then Dias += 1

            End If

            If f = Fecha Then Exit For

            f = f.AddDays(1)
        Next

        Return Dias

    End Function
    Public Function ObtenerSiguienteDiaHabil(ByVal Desde As Date) As Date
        Dim EsFinde As Boolean
        Dim EsFeriado As Boolean

        Do
            If Desde = Date.Today Then Desde = Desde.AddDays(1)
            If EsFeriado Or EsFinde Then Desde = Desde.AddDays(1)

            EsFinde = False
            EsFeriado = False

            If Desde.DayOfWeek = DayOfWeek.Saturday Then EsFinde = True
            If Desde.DayOfWeek = DayOfWeek.Sunday Then EsFinde = True
            If Existe(Desde) Then EsFeriado = True
            'EsFeriado = True
        Loop While EsFeriado Or EsFinde

        Return Desde

    End Function
    Public Function DiferenciaDiasHabiles(ByVal Desde As Date, ByVal Hasta As Date) As Integer
        Dim i As Integer = 0
        Dim j As Integer = 0

        j = CInt(IIf(Desde < Hasta, 1, -1))

        Do Until (Desde = Hasta)
            Desde = Desde.AddDays(j)

            If Not (Desde.DayOfWeek = DayOfWeek.Saturday Or Desde.DayOfWeek = DayOfWeek.Sunday Or Existe(Desde)) Then
                i += 1
            End If

        Loop

        Return i

    End Function

End Class