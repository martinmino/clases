Imports System.Data.OracleClient

Public Class Puesto
    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
        Adaptadores()
    End Sub

    Private Sub Adaptadores()
        Dim Sql As String

        Sql = "SELECT * FROM xpuestos WHERE regid_0 = :regid_0"
        da = New OracleDataAdapter(Sql, cn)

        Sql = "UPDATE xpuestos "
        Sql &= "SET sector_0 = :sector_0, "
        Sql &= "    puesto_0 = :puesto_0, "
        Sql &= "    tipo_0 = :tipo_0, "
        Sql &= "    baliza_0 = :baliza_0, "
        Sql &= "    soporte_0 = :soporte_0, "
        Sql &= "    gabinete_0 = :gabinete_0, "
        Sql &= "    cartel_0 = :cartel_0, "
        Sql &= "    tarjeta_0 = :tarjeta_0, "
        Sql &= "    lanza_0 = :lanza_0, "
        Sql &= "    pico_0 = :pico_0, "
        Sql &= "    llave_0 = :llave_0, "
        Sql &= "    vidrio_0 = :vidrio_0, "
        Sql &= "    valvula_0 = :valvula_0, "
        Sql &= "    macnum_0 = :macnum_0, "
        Sql &= "    obs_0 = :obs_0, "
        Sql &= "    estado_0 = :estado_0, "
        Sql &= "    fecha_0 = :fecha_0  "
        Sql &= "WHERE regid_0 = :regid_0"
        da.UpdateCommand = New OracleCommand(Sql, cn)

        Sql = "INSERT INTO xpuestos "
        Sql &= "VALUES (:regid_0, :sector_0, :puesto_0, :tipo_0, :baliza_0, :soporte_0, :gabinete_0, :cartel_0, :tarjeta_0, "
        Sql &= "        :lanza_0, :pico_0, :llave_0, :vidrio_0, :valvula_0, :macnum_0, :estado_0, :obs_0, :fecha_0)"
        da.InsertCommand = New OracleCommand(Sql, cn)

        With da
            .SelectCommand.Parameters.Add("regid_0", OracleType.VarChar)

            Parametro(.UpdateCommand, "regid_0", OracleType.Number, DataRowVersion.Original)
            Parametro(.UpdateCommand, "sector_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.UpdateCommand, "puesto_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "tipo_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "baliza_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "soporte_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "gabinete_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "cartel_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "tarjeta_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "lanza_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "pico_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "llave_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "vidrio_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "valvula_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "macnum_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "obs_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "estado_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.UpdateCommand, "fecha_0", OracleType.DateTime, DataRowVersion.Current)

            Parametro(.InsertCommand, "regid_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "sector_0", OracleType.Number, DataRowVersion.Current)
            Parametro(.InsertCommand, "puesto_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "tipo_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "baliza_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "soporte_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "gabinete_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "cartel_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "tarjeta_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "lanza_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "pico_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "llave_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "vidrio_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "valvula_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "macnum_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "obs_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "estado_0", OracleType.VarChar, DataRowVersion.Current)
            Parametro(.InsertCommand, "fecha_0", OracleType.DateTime, DataRowVersion.Current)

        End With

    End Sub

    Public Sub Nuevo()
        Dim dr As DataRow

        If dt Is Nothing Then
            Dim da As New OracleDataAdapter("SELECT * FROM xpuestos WHERE regid_0 = -1", cn)

            dt = New DataTable

            da.FillSchema(dt, SchemaType.Mapped)
            da.Dispose()

        Else
            dt.Clear()

        End If

        dr = dt.NewRow
        dr("regid_0") = 0
        dr("sector_0") = 0
        dr("puesto_0") = " "
        dr("tipo_0") = 0
        dr("baliza_0") = " "
        dr("soporte_0") = " "
        dr("gabinete_0") = " "
        dr("cartel_0") = " "
        dr("tarjeta_0") = " "
        dr("lanza_0") = " "
        dr("pico_0") = " "
        dr("llave_0") = " "
        dr("vidrio_0") = " "
        dr("valvula_0") = " "
        dr("macnum_0") = " "
        dr("obs_0") = " "
        dr("estado_0") = " "
        dr("fecha_0") = #12/31/1599#
        dt.Rows.Add(dr)

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
        Dim Serie As String

        'Si es nuevo obtengo nuevo numero de serie
        If dt.Rows(0).RowState = DataRowState.Added Then

            With dt.Rows(0)
                .BeginEdit()
                .Item("regid_0") = NuevoID()
                .EndEdit()
            End With

        End If

        'Obtengo el numero de serie el parque ubicado en el puesto
        Serie = dt.Rows(0).Item("macnum_0").ToString
        If Serie.Trim <> "" Then
            LimpiarPuesto(Serie)
        End If

        da.Update(dt)

    End Sub
    Private Sub LimpiarPuesto(ByVal Serie As String)
        'Quita el numero de serie de todos los puestos donde este
        Dim cm As OracleCommand
        Dim sql As String

        sql = "UPDATE xpuestos SET macnum_0 = ' ' WHERE macnum_0 = :mac"
        cm = New OracleCommand(sql, cn)
        cm.Parameters.Add("mac", OracleType.VarChar).Value = Serie
        Try
            cm.ExecuteNonQuery()
        Catch ex As Exception
        End Try
        cm.Dispose()

    End Sub
    Private Function NuevoID() As Long
        Dim da As New OracleDataAdapter("SELECT MAX(regid_0) FROM xpuestos", cn)
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
            Dim dr As DataRow = dt.Rows(0)
            Return CInt(dr("regid_0"))
        End Get
    End Property
    Public Property Sector() As Long
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CLng(dr("sector_0"))
        End Get
        Set(ByVal value As Long)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("sector_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Puesto() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("puesto_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("puesto_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Tipo() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("tipo_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("tipo_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Baliza() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("baliza_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("baliza_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Soporte() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("soporte_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("soporte_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Gabinete() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("gabinete_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("gabinete_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Cartel() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("cartel_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("cartel_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Tarjeta() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("tarjeta_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("tarjeta_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Lanza() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("lanza_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("lanza_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Pico() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("pico_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("pico_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Llave() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("llave_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("llave_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Vidrio() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("vidrio_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("vidrio_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Valvula() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("valvula_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("valvula_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Serie() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("macnum_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("macnum_0") = value
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
    Public Property Estado() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("estado_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("estado_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Fecha() As Date
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CDate(dr("fecha_0"))
        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("fecha_0") = value
            dr.EndEdit()
        End Set
    End Property

End Class