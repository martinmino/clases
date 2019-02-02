Imports System.Data.OracleClient
Imports System.Windows.Forms

Public Class Carrito
    Private Const CAPACIDAD_COURIER As Integer = 5 'Cantidad maxima de equipos para la courier
    Private Const CAPACIDAD_TOTAL_CAMIONETA As Integer = 45
    Private Const CAPACIDAD_TOTAL_COURIER As Integer = 20

    Private cn As OracleConnection
    Private dt1 As DataTable 'Tabla que contiene cantidad de equipos por dia en Camioneta
    Private dt2 As DataTable 'Tabla que contiene cantidad de equipos por dia en Corrier
    Private dt3 As DataTable 'Tabla que contiene cantidad de itn por dia en Camioneta
    Private dt4 As DataTable 'Tabla que contiene cantidad de itn por dia en Courrier
    Private dt5 As DataTable
    Private Desde As Date
    Private Hasta As Date
    Private Cant As Integer
    Private Barrio As Barrios
    Private mth As MonthCalendar = Nothing
    Private f1 As New ArrayList 'Fechas con cupo de camioneta
    Private f2 As New ArrayList 'Fechas con cipo de coutier
    Private f() As Date  'Fechas con cupo

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
        Barrio = New Barrios(cn)
    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal Barrio As Integer, ByVal mth As MonthCalendar)
        Me.New(cn)
        Me.Abrir(Barrio, mth)
    End Sub
    Public Sub Abrir(ByVal Barrio As Integer, ByVal mth As MonthCalendar)
        Me.Desde = mth.MinDate
        Me.Hasta = mth.MaxDate
        Me.mth = mth
        Me.Barrio.Abrir(Barrio)

        f = Nothing

        With mth
            .MinDate = Desde
            .MaxDate = Hasta
            .MaxSelectionCount = 1
        End With

        'Obtengo equipos por dias para camionetas
        EquiposPorDia(1, dt1)
        'Obtengo equipos por dias para couttier
        EquiposPorDia(2, dt2)
        'Obtengo intervenciones por dias para camionetas
        IntervencionesPorDia(1, dt3)
        'Obtengo intervenciones por dias para courier
        IntervencionesPorDia(2, dt4)
        'Obtengo dias de recorrido de la coutier
        DiasCourrier(dt5)

    End Sub
    Public Sub Cupo(ByVal Cant As Integer)
        Dim fe As New Feriados(cn)
        Dim DiaActual As Date

        f1.Clear()
        f2.Clear()
        f = Nothing

        DiaActual = Desde.AddDays(-1)

        While DiaActual <= Hasta
            DiaActual = DiaActual.AddDays(1)

            'Si es sabado, domingo o feriado salto
            If DiaActual.DayOfWeek = DayOfWeek.Saturday Then Continue While
            If DiaActual.DayOfWeek = DayOfWeek.Sunday Then Continue While
            If fe.Existe(DiaActual) Then Continue While

            If Cant <= CAPACIDAD_COURIER Then
                If HayCupo(DiaActual, Cant, 2) Then
                    AgregarDia(DiaActual)
                    f2.Add(DiaActual)
                    Continue While
                End If
            End If

            If HayCupo(DiaActual, Cant, 1) Then
                AgregarDia(DiaActual)
                f1.Add(DiaActual)
            End If

        End While

        mth.BoldedDates = f
        mth.UpdateBoldedDates()

    End Sub
    Private Sub AgregarDia(ByVal Dia As Date)
        Dim i As Integer = 0

        If f Is Nothing Then
            ReDim f(0)
        Else
            ReDim Preserve f(f.Length)
            i = f.Length - 1
        End If

        f(i) = Dia
    End Sub
    Private Sub EquiposPorDia(ByVal Tipo As Integer, ByRef dt As DataTable)
        Dim Sql As String
        Dim da As OracleDataAdapter

        Sql = "select xcardat_0, sum(tqty_0) as cant "
        Sql &= "from interven itn inner join yitndet yit on (itn.num_0 = yit.num_0) "
        Sql &= "where yit.typlig_0 = 1 and "
        Sql &= "      xcarzon_0 = :zona and "
        Sql &= "      xcardat_0 between :ini and :fin and "
        Sql &= "      xcartyp_0 = :modo "
        Sql &= "group by xcardat_0 "
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("ini", OracleType.DateTime).Value = Desde
        da.SelectCommand.Parameters.Add("fin", OracleType.DateTime).Value = Hasta
        da.SelectCommand.Parameters.Add("modo", OracleType.Number).Value = Tipo

        If Tipo = 1 Then
            da.SelectCommand.Parameters.Add("zona", OracleType.Number).Value = Barrio.ZonaCamioneta
        Else
            da.SelectCommand.Parameters.Add("zona", OracleType.Number).Value = Barrio.ZonaCourier
        End If

        If dt Is Nothing Then
            dt = New DataTable
        Else
            dt.Clear()
        End If

        da.Fill(dt)
        da.Dispose()
        da = Nothing

    End Sub
    Private Sub IntervencionesPorDia(ByVal Tipo As Integer, ByRef dt As DataTable)
        Dim Sql As String
        Dim da As OracleDataAdapter

        'Consulto la cantidad de intervenciones por dia
        Sql = "select xcardat_0, count(*) as cant "
        Sql &= "from interven "
        Sql &= "where xcarzon_0 = :zona and "
        Sql &= "      xcardat_0 between :ini and :fin and "
        Sql &= "      xcartyp_0 = :modo "
        Sql &= "group by xcardat_0 "
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("ini", OracleType.DateTime).Value = Desde
        da.SelectCommand.Parameters.Add("fin", OracleType.DateTime).Value = Hasta
        da.SelectCommand.Parameters.Add("modo", OracleType.Number).Value = Tipo
        If Tipo = 1 Then
            da.SelectCommand.Parameters.Add("zona", OracleType.Number).Value = Barrio.ZonaCamioneta
        Else
            da.SelectCommand.Parameters.Add("zona", OracleType.Number).Value = Barrio.ZonaCourier
        End If

        If dt Is Nothing Then
            dt = New DataTable
        Else
            dt.Clear()
        End If
        da.Fill(dt)
        da.Dispose()
        da = Nothing
    End Sub
    Private Sub DiasCourrier(ByRef dt As DataTable)
        Dim Sql As String
        Dim da As OracleDataAdapter

        'Recupero la agenda de recorridos de la courier
        Sql = "SELECT * FROM zzonasd WHERE zona_0 = :zona ORDER BY fecha_0"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("zona", OracleType.Number).Value = Barrio.ZonaCourier

        If dt Is Nothing Then
            dt = New DataTable
        Else
            dt.Clear()

        End If
        da.Fill(dt)

    End Sub

    Private Function HayCupo(ByVal Fecha As Date, ByVal Cant As Integer, ByVal Modo As Integer) As Boolean
        Dim flg As Boolean = True
        Dim i As Integer = 0
        Dim dtc As DataTable
        Dim dti As DataTable

        If Modo = 1 Then
            If Cant > CAPACIDAD_TOTAL_CAMIONETA Then flg = False
            dtc = dt1
            dti = dt3
        Else
            If Cant > CAPACIDAD_TOTAL_COURIER Then flg = False
            If Not EsDiaCourier(Fecha) Then flg = False
            dtc = dt2
            dti = dt4
        End If

        'Busco la fecha
        For Each dr As DataRow In dtc.Rows
            If CDate(dr("xcardat_0")) = Fecha Then
                i = CInt(dr("cant")) + Cant

                If Modo = 1 Then
                    If i > CAPACIDAD_TOTAL_CAMIONETA Then flg = False
                Else
                    If i > CAPACIDAD_TOTAL_COURIER Then flg = False
                End If

                'Consulto si se llego al tope de intervenciones para ese dia
                For Each dri As DataRow In dti.Rows
                    If CDate(dri("xcardat_0")) = Fecha Then
                        If CInt(dri("cant")) >= 8 Then flg = False
                        Exit For
                    End If
                Next

            End If

        Next

        dtc = Nothing
        dti = Nothing

        Return flg

    End Function
    
    Private Function EsDiaCourier(ByVal Fecha As Date) As Boolean
        Dim dr As DataRow = Nothing
        Dim flg As Boolean = False

        For Each dr In dt5.Rows
            If CDate(dr("fecha_0")) >= Fecha Then Exit For
        Next

        If dr IsNot Nothing Then

            Select Case Fecha.DayOfWeek
                Case DayOfWeek.Monday
                    flg = CInt(dr("uvyday1_0")) = 2

                Case DayOfWeek.Tuesday
                    flg = CInt(dr("uvyday2_0")) = 2

                Case DayOfWeek.Wednesday
                    flg = CInt(dr("uvyday3_0")) = 2

                Case DayOfWeek.Thursday
                    flg = CInt(dr("uvyday4_0")) = 2

                Case DayOfWeek.Friday
                    flg = CInt(dr("uvyday5_0")) = 2

                Case Else
                    flg = False

            End Select

        End If

        Return flg

    End Function
    Public Sub ObtenerIntervenciones(ByVal Zona As Integer, ByVal Fecha As Date, ByVal Tipo As Integer, ByVal dgv As DataGridView)
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As DataTable

        If dgv.DataSource Is Nothing Then
            dt = New DataTable
            dgv.DataSource = dt
        Else
            dt = CType(dgv.DataSource, DataTable)
        End If

        Sql = "SELECT itn.num_0, itn.xcardat_0, sum(tqty_0) as cant "
        Sql &= "FROM interven itn INNER JOIN yitndet yit on (itn.num_0 = yit.num_0) "
        Sql &= "WHERE typlig_0 = 1 AND "
        Sql &= "      tripnum_0 = ' ' AND "
        Sql &= "      zflgtrip_0 IN (1, 6) AND "
        Sql &= "      xcarzon_0 = :xcarzon AND "
        Sql &= "      xcardat_0 <= :xcardat AND "
        Sql &= "      xcartyp_0 = :xcartyp "
        Sql &= "GROUP BY itn.num_0, itn.xcardat_0"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("xcarzon", OracleType.Number).Value = Zona
        da.SelectCommand.Parameters.Add("xcardat", OracleType.DateTime).Value = Fecha
        da.SelectCommand.Parameters.Add("xcartyp", OracleType.Number).Value = Tipo

        dt.Clear()
        da.Fill(dt)
        da.Dispose()

    End Sub

    Public ReadOnly Property Zona(ByVal Fecha As Date) As Integer
        Get
            Select Case Tipo(Fecha)
                Case 1
                    Return Barrio.ZonaCamioneta
                Case 2
                    Return Barrio.ZonaCourier
                Case Else
                    Return 0
            End Select
        End Get
    End Property
    Public ReadOnly Property Tipo(ByVal Fecha As Date) As Integer
        Get
            If f1.IndexOf(Fecha) > -1 Then
                Return 1
            ElseIf f2.IndexOf(Fecha) > -1 Then
                Return 2
            Else
                Return 0
            End If
        End Get
    End Property

End Class