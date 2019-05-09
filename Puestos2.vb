﻿Imports System.Data.OracleClient
Imports System.ComponentModel

Public Class Puestos2
    Inherits BindingList(Of Puesto2)

    Private cn As OracleConnection

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
    End Sub
    Public Function Abrir(ByVal idSector As Integer) As Boolean
        Dim da As OracleDataAdapter
        Dim Sql As String
        Dim dt As New DataTable

        Sql = "SELECT * FROM xpuestos2 WHERE idsector_0 = :idsector"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("idsector", OracleType.Number).Value = idSector
        da.Fill(dt)

        ArmarColeccion(dt)

    End Function
    Private Sub ArmarColeccion(ByVal dt As DataTable)
        Me.Clear()

        For Each dr As DataRow In dt.Rows
            Dim p As New Puesto2(cn)
            If p.Abrir(dr) Then Me.Add(p)
        Next
    End Sub
    Public Sub Grabar()
        For Each p As Puesto2 In Me
            p.Grabar()
        Next
    End Sub

End Class
