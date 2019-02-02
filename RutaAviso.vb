Imports System.Data.OracleClient

Public Class RutaAviso
    Private da As OracleDataAdapter
    Private dt As New DataTable
    Private cn As OracleConnection

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String
        Me.cn = cn

        Sql = "SELECT * "
        Sql &= "FROM xrutax "
        Sql &= "WHERE ruta_0 = :ruta and cliente_0 = :cli AND suc_0 = :suc "

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("ruta", OracleType.VarChar)
        da.SelectCommand.Parameters.Add("cli", OracleType.VarChar)
        da.SelectCommand.Parameters.Add("suc", OracleType.VarChar)

        da.InsertCommand = New OracleCommandBuilder(da).GetInsertCommand
        da.DeleteCommand = New OracleCommandBuilder(da).GetDeleteCommand
        da.UpdateCommand = New OracleCommandBuilder(da).GetUpdateCommand

    End Sub
    Public Sub Abrir(ByVal ruta As String, ByVal cli As String, ByVal suc As String)
        da.SelectCommand.Parameters("ruta").Value = ruta
        da.SelectCommand.Parameters("cli").Value = cli
        da.SelectCommand.Parameters("suc").Value = suc

        Try
            dt.Clear()
            da.Fill(dt)

            If dt.Rows.Count = 0 Then
                Dim dr As DataRow
                dr = dt.NewRow
                dr("ruta_0") = ruta
                dr("cliente_0") = cli
                dr("suc_0") = suc
                dr("confirma_0") = 0
                dr("modo_0") = 0
                dr("obs_0") = " "
                dr("fecha_0") = #12/31/1599#
                dr("hora_0") = "0000"
                dt.Rows.Add(dr)
            End If

        Catch ex As Exception
        End Try

    End Sub
    Public Sub Grabar()
        Try
            Dim dr As DataRow = dt.Rows(0)

            dr.BeginEdit()
            dr("fecha_0") = Date.Today
            dr("hora_0") = Now.ToString("HHmm")
            dr.EndEdit()

            da.Update(dt)

        Catch ex As Exception
        End Try
    End Sub

    Public Property Ruta() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CInt(dr("ruta_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("ruta_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Cliente() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("cliente_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("cliente_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Sucursal() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("suc_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("suc_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Confirma() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CInt(dr("confirma_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("confirma_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Modo() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CInt(dr("modo_0"))
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("modo_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Observaciones() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("obs_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("obs_0") = value
            dr.EndEdit()
        End Set
    End Property
End Class
