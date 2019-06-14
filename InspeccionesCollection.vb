Imports System.Data.OracleClient
Imports System.ComponentModel

Public Class InspeccionesCollection
    Inherits BindingList(Of Inspeccion)

    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Public dt As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String

        Me.cn = cn

        Sql = "SELECT * " _
            & "FROM xinspeccio " _
            & "WHERE itn_0 = :itn"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("itn", OracleType.VarChar)
        da.UpdateCommand = New OracleCommandBuilder(da).GetUpdateCommand
        da.InsertCommand = New OracleCommandBuilder(da).GetInsertCommand
        da.DeleteCommand = New OracleCommandBuilder(da).GetDeleteCommand
    End Sub
    Public Sub Abrir(ByVal Intervencion As String)

        da.SelectCommand.Parameters("itn").Value = Intervencion

        Me.Clear()
        dt.Clear()

        da.Fill(dt)

        ArmarColeccion(dt)

    End Sub
    Private Sub ArmarColeccion(ByVal dt As DataTable)
        Me.Clear()

        For Each dr As DataRow In dt.Rows
            Dim c As New Inspeccion(cn)
            If c.Abrir(dr) Then Me.Add(c)
        Next
    End Sub
    Public Function Buscar(ByVal id As Integer) As Inspeccion
        Dim i As Inspeccion

        For Each i In Me
            If i.id = id Then Return i
        Next

        Return Nothing

    End Function
    Public Sub Grabar()
        da.Update(dt)
    End Sub
End Class