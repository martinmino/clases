Imports System.Data.OracleClient

Public Class Segto2
    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable

    Sub New(ByVal cn As OracleConnection)

        Me.cn = cn
        Adaptadores()

    End Sub

    Private Sub Adaptadores()
        da = New OracleDataAdapter("SELECT * FROM xsegto2 WHERE itn_0 = :itn_0", cn)
        da.SelectCommand.Parameters.Add("itn_0", OracleType.VarChar)
        da.UpdateCommand = New OracleCommandBuilder(da).GetUpdateCommand
        da.InsertCommand = New OracleCommandBuilder(da).GetInsertCommand
        da.DeleteCommand = New OracleCommandBuilder(da).GetDeleteCommand
    End Sub
    Public Function Abrir(ByVal itn As String) As Boolean
        dt.Clear()

        da.SelectCommand.Parameters("itn_0").Value = itn.ToUpper

        Try
            da.Fill(dt)

            Abrir = (dt.Rows.Count = 1)

        Catch ex As Exception
            Abrir = False

        End Try

    End Function
    Public Function Nuevo(ByVal itn As String) As Boolean
        Dim dr As DataRow

        'Busco si ya existe la intervencion
        da.SelectCommand.Parameters("itn_0").Value = itn.ToUpper
        dt.Clear()

        Try
            da.Fill(dt)

            If dt.Rows.Count = 0 Then
                dr = dt.NewRow
                dr("itn_0") = itn
                dr("dat_0") = #12/31/1599#
                dr("dat_1") = #12/31/1599#
                dr("dat_2") = #12/31/1599#
                dr("dat_3") = #12/31/1599#
                dr("dat_4") = #12/31/1599#
                dr("equipos_0") = 0
                dr("cant_0") = 0
                dr("cant_1") = 0
                dr("rech_0") = 0
                dr("rech_1") = 0
                dt.Rows.Add(dr)
                Nuevo = True

            Else
                Nuevo = False

            End If

        Catch ex As Exception
            Nuevo = False
            Exit Function

        End Try

    End Function
    Public Function Grabar() As Boolean
        Try
            da.Update(dt)
            Return True

        Catch ex As Exception
            Return False

        End Try

    End Function
    Public Property Equipos() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CType(dr("equipos_0"), Integer)
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("equipos_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property EquiposLeidos() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CType(dr("cant_0"), Integer)
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("cant_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property EquiposRechazados() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CType(dr("rech_0"), Integer)
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("rech_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property MangasLeidas() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CType(dr("cant_1"), Integer)
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("cant_1") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property MangasRechazadas() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CType(dr("rech_1"), Integer)
        End Get
        Set(ByVal value As Integer)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("rech_1") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property FechaIngresoPlanta() As Date
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CType(dr("dat_0"), Date)
        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("dat_0") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property FechaIngresoService() As Date
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CType(dr("dat_1"), Date)
        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("dat_1") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property FechaIngresoAdministracion() As Date
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CType(dr("dat_2"), Date)
        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("dat_2") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property FechaIngresoLogistica() As Date
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CType(dr("dat_3"), Date)
        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("dat_3") = value
            dr.EndEdit()
        End Set
    End Property
    Public Property FechaEntregado() As Date
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CType(dr("dat_4"), Date)
        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("dat_4") = value
            dr.EndEdit()
        End Set
    End Property

End Class