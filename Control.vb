Imports System.Data.OracleClient

Public Class Control

    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As DataTable
    Private dr As DataRow

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String

        Me.cn = cn

        Sql = "SELECT * " _
            & "FROM xcontroles " _
            & "WHERE itn_0 = :itn"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("itn", OracleType.VarChar)
        da.InsertCommand = New OracleCommandBuilder(da).GetInsertCommand
        da.UpdateCommand = New OracleCommandBuilder(da).GetUpdateCommand
        da.DeleteCommand = New OracleCommandBuilder(da).GetDeleteCommand

    End Sub
    Public Function Abrir(ByVal Itn As String) As Boolean
        If dt Is Nothing Then dt = New DataTable

        dt.Clear()
        da.SelectCommand.Parameters("itn").Value = Itn
        da.Fill(dt)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            Return True
        End If

        Return False

    End Function
    Public Sub Nuevo()
        If dt Is Nothing Then
            dt = New DataTable
            da.SelectCommand.Parameters("itn").Value = "xxx"
            da.Fill(dt)
        End If

        dt.Clear()
        dr = dt.NewRow
        dr("itn_0") = " "
        dr("dat_0") = #12/31/1599#
        dr("bpcnum_0") = " "
        dr("bpaadd_0") = " "
        dr("estado_0") = 1
        dt.Rows.Add(dr)
    End Sub
    Public Sub Borrar()
        dr.Delete()
    End Sub
    Public Sub Grabar()
        da.Update(dt)
    End Sub
    Public Function Inspecciones() As InspeccionesCollection
        Dim i As New InspeccionesCollection(cn)
        i.Abrir(Me.id)
        Return i
    End Function
    Public Property id() As String
        Get
            Return dr("itn_0").ToString
        End Get
        Set(ByVal value As String)
            dr.BeginEdit()
            dr("itn_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Fecha() As Date
        Get
            Return CDate(dr("dat_0"))
        End Get
        Set(ByVal value As Date)
            dr.BeginEdit()
            dr("dat_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Cliente() As String
        Get
            Return dr("bpcnum_0").ToString
        End Get
        Set(ByVal value As String)
            dr.BeginEdit()
            dr("bpcnum_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Sucursal() As String
        Get
            Return dr("bpaadd_0").ToString
        End Get
        Set(ByVal value As String)
            dr.BeginEdit()
            dr("bpaadd_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Estado() As Integer
        Get
            Return CInt(dr("estado_0"))
        End Get
        Set(ByVal value As Integer)
            dr.BeginEdit()
            dr("estado_0") = value
            dr.EndEdit()
        End Set
    End Property

End Class