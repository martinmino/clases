Imports System.Data.OracleClient

Public Class Foto
    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As DataTable = Nothing
    Private dr As DataRow = Nothing

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String

        Me.cn = cn

        Sql = "SELECT * FROM xfotos WHERE id_0 = :id "
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("id", OracleType.Number)
        da.InsertCommand = New OracleCommandBuilder(da).GetInsertCommand
        da.UpdateCommand = New OracleCommandBuilder(da).GetUpdateCommand
        da.DeleteCommand = New OracleCommandBuilder(da).GetDeleteCommand
    End Sub
    Public Sub Abrir(ByVal id As Integer)
        If dt Is Nothing Then dt = New DataTable
        dt.Clear()
        da.SelectCommand.Parameters("id").Value = id
        da.Fill(dt)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
        Else
            dr = Nothing
        End If

    End Sub
    Public Sub Nuevo(ByVal id As Integer)
        Abrir(id)

        If dr Is Nothing Then
            dr = dt.NewRow
            dr("id_0") = id
            dt.Rows.Add(dr)
        End If

    End Sub
    Public Sub Grabar()
        da.Update(dt)
    End Sub

    Public Property Id() As Integer
        Get
            Return CInt(dr("id_0"))
        End Get
        Set(ByVal value As Integer)
            dr.BeginEdit()
            dr("id_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Descripcion() As String
        Get
            Return dr("des_0").ToString
        End Get
        Set(ByVal value As String)
            dr.BeginEdit()
            dr("des_0") = IIf(value.Trim = "", " ", value)
            dr.EndEdit()
        End Set
    End Property
    Public Property Inspeccion() As Integer
        Get
            Return CInt(dr("inspeccion_0"))
        End Get
        Set(ByVal value As Integer)
            dr.BeginEdit()
            dr("inspeccion_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Foto() As String
        Get
            Return dr("foto_0").ToString
        End Get
        Set(ByVal value As String)
            dr.BeginEdit()
            dr("foto_0") = value
            dr.EndEdit()
        End Set
    End Property

End Class