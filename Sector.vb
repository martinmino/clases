Imports System.Data.OracleClient

Public Class Sector
    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
        Adaptadores()
    End Sub

    Private Sub Adaptadores()
        Dim Sql As String

        Sql = "SELECT * FROM xsectores WHERE regid_0 = :regid_0"
        da = New OracleDataAdapter(Sql, cn)

        Sql = "UPDATE xsectores "
        Sql &= "SET bpcnum_0 = :bpcnum_0, "
        Sql &= "    fcyitn_0 = :fcyitn_0, "
        Sql &= "    nombre_0 = :nombre_0 "
        Sql &= "WHERE regid_0 = :regid_0"
        da.UpdateCommand = New OracleCommand(Sql, cn)

        Sql = "INSERT INTO xsectores VALUES (:regid_0, :bpcnum_0, :fcyitn_0, :nombre_0)"
        da.InsertCommand = New OracleCommand(Sql, cn)

        With da
            .SelectCommand.Parameters.Add("regid_0", OracleType.VarChar)

            Parametro(.UpdateCommand, "regid_0", OracleType.Number, DataRowVersion.Original)
            Parametro(.UpdateCommand, "bpcnum_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "fcyitn_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "nombre_0", OracleType.VarChar, DataRowVersion.Current)

            Parametro(.InsertCommand, "regid_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "bpcnum_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "fcyitn_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "nombre_0", OracleType.VarChar, DataRowVersion.Current)

        End With

    End Sub

    Public Sub Nuevo()
        Dim dr As DataRow

        If dt Is Nothing Then
            Dim da As New OracleDataAdapter("SELECT * FROM xsectores WHERE regid_0 = -1", cn)

            dt = New DataTable

            da.FillSchema(dt, SchemaType.Mapped)
            da.Dispose()

        Else
            dt.Clear()

        End If

        dr = dt.NewRow
        dr("regid_0") = 0
        dr("bpcnum_0") = " "
        dr("fcyitn_0") = " "
        dr("nombre_0") = " "
        dt.Rows.Add(dr)

    End Sub
    Public Sub Nuevo(ByVal Cliente As String, ByVal Sucursal As String, ByVal Nombre As String)
        Dim dr As DataRow

        Nuevo()

        dr = dt.Rows(0)

        dr.BeginEdit()
        dr("regid_0") = 0
        dr("bpcnum_0") = Cliente
        dr("fcyitn_0") = Sucursal
        dr("nombre_0") = Nombre
        dr.EndEdit()

    End Sub
    Public Function Abrir(ByVal ID As Long) As Boolean
        da.SelectCommand.Parameters("regid_0").Value = ID

        If dt Is Nothing Then
            dt = New DataTable

        Else
            dt.Clear()

        End If

        Try
            da.Fill(dt)

            Return dt.Rows.Count = 1

        Catch ex As Exception

            Return False

        End Try

    End Function
    Public Sub Grabar()
        'Si es nuevo obtengo nuevo numero de serie
        If dt.Rows(0).RowState = DataRowState.Added Then

            With dt.Rows(0)
                .BeginEdit()
                .Item("regid_0") = NuevoID
                .EndEdit()
            End With

        End If

        da.Update(dt)

    End Sub
    Public Function Buscar(ByVal Nombre As String, ByVal Cliente As String, ByVal Sucursal As String) As Long
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim dr As DataRow

        Sql = "SELECT * FROM xsectores WHERE nombre_0 = :nombre_0 AND bpcnum_0 = :bpcnum_0 AND fcyitn_0 = :fcyitn_0"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("nombre_0", OracleType.VarChar).Value = Nombre
        da.SelectCommand.Parameters.Add("bpcnum_0", OracleType.VarChar).Value = Cliente
        da.SelectCommand.Parameters.Add("fcyitn_0", OracleType.VarChar).Value = Sucursal

        da.Fill(dt)
        da.Dispose()

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            Buscar = CLng(dr("regid_0"))

            Abrir(CLng(dr("regid_0")))

        Else
            Buscar = 0

        End If

        dt.Dispose()

    End Function
    Private Function NuevoID() As Long
        Dim da As New OracleDataAdapter("SELECT MAX(regid_0) FROM xsectores", cn)
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

    Public ReadOnly Property ID() As Long
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CLng(dr("regid_0"))
        End Get
    End Property
    Public Property Cliente() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("bpcnum_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("bpcnum_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Sucursal() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("fcyitn_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("fcyitn_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Nombre() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("nombre_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("nombre_0") = value
            dr.EndEdit()
        End Set
    End Property

End Class
