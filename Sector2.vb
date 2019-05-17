Imports System.Data.OracleClient

Public Class Sector2
    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As DataTable
    Private dr As DataRow

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String

        Me.cn = cn

        Sql = "SELECT * " _
            & "FROM xsectores2 " _
            & "WHERE id_0 = :id"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("id", OracleType.Number)
        da.InsertCommand = New OracleCommandBuilder(da).GetInsertCommand
        da.UpdateCommand = New OracleCommandBuilder(da).GetUpdateCommand
        da.DeleteCommand = New OracleCommandBuilder(da).GetDeleteCommand

    End Sub
    Public Sub Nuevo(ByVal Cliente As String, ByVal Sucursal As String)
        If dt Is Nothing Then
            dt = New DataTable
            da.SelectCommand.Parameters("id").Value = -1
            da.Fill(dt)
        End If
        dt.Clear()
        dr = dt.NewRow
        dr("id_0") = 0
        dr("numero_0") = " "
        dr("sector_0") = " "
        dr("bpcnum_0") = Cliente
        dr("fcyitn_0") = Sucursal
        dt.Rows.Add(dr)
    End Sub
    Friend Function Abrir(ByVal dr2 As DataRow) As Boolean
        dt = Nothing
        dt = dr2.Table.Clone

        dr = dt.NewRow

        For i = 0 To dt.Columns.Count - 1
            Me.dr(i) = dr2(i)
        Next

        dt.Rows.Add(Me.dr)
        dt.AcceptChanges()

        Return True
    End Function
    Public Function Abrir(ByVal id As Integer) As Boolean
        If dt Is Nothing Then dt = New DataTable
        dt.Clear()
        da.SelectCommand.Parameters("id").Value = id
        da.Fill(dt)
        Return dt.Rows.Count > 0
    End Function
    Public Function Puestos() As Puestos2
        Dim p As New Puestos2(cn)
        p.Abrir(Me.id)
        Return p
    End Function
    Public Sub Eliminar()
        dr.Delete()
    End Sub
    Public Sub Grabar()
        If dr.RowState <> DataRowState.Deleted Then
            If CInt(dr("id_0")) = 0 Then
                dr.BeginEdit()
                dr("id_0") = SiguienteId()
                dr.EndEdit()
            End If
        End If

        da.Update(dt)
    End Sub
    Private Function SiguienteId() As Long
        Dim da As New OracleDataAdapter("SELECT MAX(id_0) FROM xsectores2", cn)
        Dim dt As New DataTable
        Dim dr As DataRow

        da.Fill(dt)
        dr = dt.Rows(0)

        If IsDBNull(dr(0)) Then
            Return 1

        Else
            Return CLng(dr(0)) + 1

        End If

    End Function

    Public ReadOnly Property Id() As Integer
        Get
            Return CInt(dr("id_0"))
        End Get
    End Property
    Public Property Numero() As String
        Get
            Return dr("numero_0").ToString
        End Get
        Set(ByVal value As String)
            dr.BeginEdit()
            dr("numero_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Nombre() As String
        Get
            Return dr("sector_0").ToString
        End Get
        Set(ByVal value As String)
            dr.BeginEdit()
            dr("sector_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property ClienteId() As String
        Get
            Return dr("bpcnum_0").ToString
        End Get
        Set(ByVal value As String)
            dr.BeginEdit()
            dr("bpcnum_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property Cliente() As Cliente
        Get
            Dim c As New Cliente(cn)
            c.Abrir(ClienteId)
            Return c
        End Get
    End Property
    Public Property SucursalId() As String
        Get
            Return dr("fcyitn_0").ToString
        End Get
        Set(ByVal value As String)
            dr.BeginEdit()
            dr("fcyitn_0") = value
            dr.EndEdit()
        End Set
    End Property

End Class
