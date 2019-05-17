Imports System.Data.OracleClient

Public Class Inspeccion
    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As DataTable
    Private dr As DataRow

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String

        Me.cn = cn

        Sql = "SELECT * " _
            & "FROM xinspeccio " _
            & "WHERE id_0 = :id"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("id", OracleType.Number)
        da.InsertCommand = New OracleCommandBuilder(da).GetInsertCommand
        da.UpdateCommand = New OracleCommandBuilder(da).GetUpdateCommand
        da.DeleteCommand = New OracleCommandBuilder(da).GetDeleteCommand

    End Sub
    Public Function Nuevo() As Boolean
        If dt Is Nothing Then
            dt = New DataTable
            da.SelectCommand.Parameters("id").Value = -1
            da.Fill(dt)
        End If

        dt.Clear()
        dr = dt.NewRow
        dr("id_0") = 0
        dr("idcontrol_0") = 0
        dr("idpuesto_0") = 0
        dr("idsector_0") = 0
        dr("tipo_0") = 0
        dr("nro_0") = " "
        dr("ubicacion_0") = " "
        dr("luz_0") = 1
        dr("cartel_0") = 1
        dr("cinta_0") = 1
        dr("equipo_0") = 1
        dr("agente_0") = " "
        dr("capacidad_0") = " "
        dr("cilindro_0") = " "
        dr("vto_0") = #12/31/1599#
        dr("vencido_0") = 1
        dr("ausente_0") = 1
        dr("obstruido_0") = 1
        dr("carro_0") = 1
        dr("usado_0") = 1
        dr("despintado_0") = 1
        dr("despresu_0") = 1
        dr("altura_0") = 1
        dr("senal_0") = 1
        dr("senalalt_0") = 1
        dr("senalbal_0") = 1
        dr("tarjeta_0") = 1
        dr("precinto_0") = 1
        dr("soporte_0") = 1
        dr("ruptura_0") = 1
        dr("manguera_0") = 1
        dr("otro_0") = 1
        dr("valvula_0") = 1
        dr("pico_0") = 1
        dr("lanza_0") = 1
        dr("vidrio_0") = 1
        dr("llave_0") = 1
        dt.Rows.Add(dr)

    End Function
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

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            Return True
        End If

        Return False

    End Function
    Public Sub Grabar()
        da.Update(dt)
    End Sub
    Public Property id() As Integer
        Get
            Return CInt(dr("id_0"))
        End Get
        Set(ByVal value As Integer)
            dr.BeginEdit()
            dr("id_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Control() As Integer
        Get
            Return CInt(dr("idcontrol_0"))
        End Get
        Set(ByVal value As Integer)
            dr.BeginEdit()
            dr("idcontrol_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Puesto() As Integer
        Get
            Return CInt(dr("idcontrol_0"))
        End Get
        Set(ByVal value As Integer)
            dr.BeginEdit()
            dr("idcontrol_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Sector() As Integer
        Get
            Return CInt(dr("idpuesto_0"))
        End Get
        Set(ByVal value As Integer)
            dr.BeginEdit()
            dr("idpuesto_0") = value
            dr.EndEdit()
        End Set
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
    Public Property Nro() As String
        Get
            Return dr("nro_0").ToString
        End Get
        Set(ByVal value As String)
            dr.BeginEdit()
            dr("nro_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Nro() As String
        Get
            Return dr("ubicacion_0").ToString
        End Get
        Set(ByVal value As String)
            dr.BeginEdit()
            dr("ubicacion_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Luz() As Boolean
        Get
            Return CBool(IIf(CInt(dr("luz_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("luz_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Cartel() As Boolean
        Get
            Return CBool(IIf(CInt(dr("cartel_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("cartel_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Cinta() As Boolean
        Get
            Return CBool(IIf(CInt(dr("cinta_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("cinta_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Equipo() As String
        Get
            Return dr("equipo_0").ToString
        End Get
        Set(ByVal value As String)
            dr.BeginEdit()
            dr("equipo_0") = value
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
    Public Property Cilindro() As String
        Get
            Return dr("cilindro_0").ToString
        End Get
        Set(ByVal value As String)
            dr.BeginEdit()
            dr("cilindro_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Vto() As Date
        Get
            Return CDate(dr("vto_0"))
        End Get
        Set(ByVal value As Date)
            dr.BeginEdit()
            dr("vto_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property Vencido() As Boolean
        Get
            Return CBool(IIf(CInt(dr("vencido_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("vencido_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Ausente() As Boolean
        Get
            Return CBool(IIf(CInt(dr("ausente_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("ausente_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Obstruido() As Boolean
        Get
            Return CBool(IIf(CInt(dr("obstruido_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("obstruido_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Carro() As Boolean
        Get
            Return CBool(IIf(CInt(dr("carro_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("carro_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Usado() As Boolean
        Get
            Return CBool(IIf(CInt(dr("usado_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("usado_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Despintado() As Boolean
        Get
            Return CBool(IIf(CInt(dr("despintado_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("despintado_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Despresurizado() As Boolean
        Get
            Return CBool(IIf(CInt(dr("despresu_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("despresu_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Altura() As Boolean
        Get
            Return CBool(IIf(CInt(dr("altura_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("altura_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Senalizacion() As Boolean
        Get
            Return CBool(IIf(CInt(dr("senal_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("senal_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property SenalizacionAltura() As Boolean
        Get
            Return CBool(IIf(CInt(dr("senalalt_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("senalalt_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property SenalizacionBaliza() As Boolean
        Get
            Return CBool(IIf(CInt(dr("senalbal_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("senalbal_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Tarjeta() As Boolean
        Get
            Return CBool(IIf(CInt(dr("tarjeta_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("tarjeta_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Precinto() As Boolean
        Get
            Return CBool(IIf(CInt(dr("precinto_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("precinto_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Soporte() As Boolean
        Get
            Return CBool(IIf(CInt(dr("soporte_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("soporte_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Ruptura() As Boolean
        Get
            Return CBool(IIf(CInt(dr("ruptura_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("ruptura_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Manguera() As Boolean
        Get
            Return CBool(IIf(CInt(dr("manguera_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("manguera_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Otro() As Boolean
        Get
            Return CBool(IIf(CInt(dr("otro_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("otro_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Valvula() As Boolean
        Get
            Return CBool(IIf(CInt(dr("valvula_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("valvula_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Pico() As Boolean
        Get
            Return CBool(IIf(CInt(dr("pico_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("pico_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Lanza() As Boolean
        Get
            Return CBool(IIf(CInt(dr("lanza_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("lanza_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Vidrio() As Boolean
        Get
            Return CBool(IIf(CInt(dr("vidrio_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("vidrio_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property
    Public Property Llave() As Boolean
        Get
            Return CBool(IIf(CInt(dr("llave_0")) = 2, True, False))
        End Get
        Set(ByVal value As Boolean)
            dr.BeginEdit()
            dr("llave_0") = IIf(value, 2, 1)
            dr.EndEdit()
        End Set
    End Property

End Class