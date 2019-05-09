Imports System.Data.OracleClient

Public Class Puesto2
    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As DataTable
    Private dr As DataRow

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
        Adaptadores()
    End Sub
    Private Sub Adaptadores()
        Dim Sql As String

        Sql = "SELECT * FROM xpuestos2 WHERE id_0 = :id"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("id", OracleType.Number)
        da.InsertCommand = New OracleCommandBuilder(da).GetInsertCommand
        da.UpdateCommand = New OracleCommandBuilder(da).GetInsertCommand
        da.DeleteCommand = New OracleCommandBuilder(da).GetDeleteCommand

        da.SelectCommand.Parameters("id").Value = -1
        da.Fill(dt)

    End Sub
    Public Function Nuevo(ByVal idSector As Integer, ByVal NroPuesto As String, ByVal Ubicacion As String, ByVal Tipo As Integer) As Boolean
        'Tipo: 1=Extintor | 2=Hidrante
        dt.Clear()
        dr = dt.NewRow
        dr("id") = 0
        dr("nropuesto_0") = " "
        dr("ubicacion_0") = " "
        dr("orden_0") = 0
        dr("idSector") = idSector
        dr("tipo") = Tipo
        dr("agente_0") = " "
        dr("capacidad_0") = " "
        dr("equipo_0") = " "
        dr("inspeccion_0") = " "
        dt.Rows.Add(dr)
    End Function
    Public Function Abrir(ByVal id As Integer) As Boolean
        dt.Clear()
        dr = Nothing

        da.SelectCommand.Parameters("id").Value = id
        da.Fill(dt)

        If dt.Rows.Count > 0 Then dr = dt.Rows(0)

        Return dt.Rows.Count > 0
    End Function
    Friend Function Abrir(ByVal dr2 As DataRow) As Boolean
        dt = Nothing
        dt = dr2.Table.Clone

        dr = dt.NewRow

        For i = 0 To dt.Columns.Count - 1
            Me.dr(i) = dr(i)
        Next

        dt.Rows.Add(Me.dr)
        dt.AcceptChanges()

        Return True
    End Function
    Public Sub Grabar()

        If CInt(dr("id")) = 0 Then
            dr.BeginEdit()
            dr("id") = SiguienteId()
            dr.EndEdit()
        End If

        da.Update(dt)
    End Sub
    Private Function SiguienteId() As Integer
        Dim da As New OracleDataAdapter("SELECT MAX(id_0) FROM xpuestos2", cn)
        Dim dt As New DataTable
        Dim dr As DataRow

        da.Fill(dt)
        dr = dt.Rows(0)

        If IsDBNull(dr(0)) Then
            Return 1

        Else
            Return CInt(dr(0)) + 1

        End If

    End Function
    Public ReadOnly Property id() As Integer
        Get
            Return CInt(dr("id"))
        End Get
    End Property
    Public Property NroPuesto() As String
        Get
            Return dr("nropuesto_0").ToString
        End Get
        Set(ByVal value As String)
            dr.BeginEdit()
            dr("nropuesto_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Ubicacion() As String
        Get
            Return dr("ubicacion_0").ToString
        End Get
        Set(ByVal value As String)
            dr.BeginEdit()
            dr("ubicacion_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Orden() As Integer
        Get
            Return CInt(dr("orden_0"))
        End Get
        Set(ByVal value As Integer)
            dr.BeginEdit()
            dr("orden_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property SectorId() As Integer
        Get
            Return CInt(dr("idsector_0"))
        End Get
        Set(ByVal value As Integer)
            dr.BeginEdit()
            dr("idsector_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property Sector() As Sector2
        Get
            Dim s As Sector2
            s = New Sector2(cn)
            s.Abrir(Me.SectorId)
            Return s
        End Get
    End Property
    Public Property Tipo() As Integer
        Get
            Return CInt(dr("tipo_0"))
        End Get
        Set(ByVal value As Integer)
            dr.BeginEdit()
            dr("tipo_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Agente() As String
        Get
            Return dr("agente_0").ToString
        End Get
        Set(ByVal value As String)
            dr.BeginEdit()
            dr("agente_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Capacidad() As String
        Get
            Return dr("capacidad_0").ToString
        End Get
        Set(ByVal value As String)
            dr.BeginEdit()
            dr("capacidad_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property EquipoId() As String
        Get
            Return dr("equipo_0").ToString
        End Get
        Set(ByVal value As String)
            dr.BeginEdit()
            dr("equipo_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property Equipo() As Parque
        Get
            Dim mac As New Parque(cn)
            mac.Abrir(Me.EquipoId)
            Return mac
        End Get
    End Property
    Public Property UltimaInspeccion() As String
        Get
            Return dr("inspeccion_0").ToString
        End Get
        Set(ByVal value As String)
            dr.BeginEdit()
            dr("inspeccion_0") = value
            dr.EndEdit()
        End Set
    End Property

End Class