Imports System.Data.OracleClient
Imports System.Windows.Forms

Public Class Sociedad

    Private cn As OracleConnection
    Private dt As New DataTable
    Private da As OracleDataAdapter

    'NEW
    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
        Dim Sql As String

        Sql = "SELECT * FROM company WHERE cpy_0 = :cpy"
        da = New OracleDataAdapter(Sql, cn)
        With da.SelectCommand.Parameters
            .Add("cpy", OracleType.VarChar)
        End With

    End Sub
    'SUB
    Public Sub abrir(ByVal Codigo As String)
        da.SelectCommand.Parameters("cpy").Value = Codigo
        dt.Clear()
        da.Fill(dt)
    End Sub
    Public Function AbrirPorCuit(ByVal cuit As String) As Boolean
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim dr As DataRow

        Sql = "SELECT * FROM company WHERE crn_0 = :crn"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("crn", OracleType.VarChar).Value = cuit
        da.Fill(dt)
        da.Dispose()

        If dt.Rows.Count = 1 Then
            dr = dt.Rows(0)
            abrir(dr("cpy_0").ToString)
            Return True

        Else
            Return False

        End If

    End Function

    'FUNCTION
    Shared Function Sociedades(ByVal cn As OracleConnection, Optional ByVal IncluirBlanco As Boolean = False) As DataTable
        Dim Sql As String = "SELECT cpy_0, cpynam_0 FROM company ORDER BY cpy_0"
        Dim da As New OracleDataAdapter(Sql, cn)
        Dim dt As New DataTable
        Dim dr As DataRow

        da.Fill(dt)

        If IncluirBlanco Then
            dr = dt.NewRow
            dr(0) = " "
            dr(1) = " "
            dt.Rows.InsertAt(dr, 0)
        End If

        Return dt

    End Function
    Shared Sub Sociedades(ByVal cn As OracleConnection, ByVal cbo As ComboBox, Optional ByVal IncluirBlanco As Boolean = False)
        With cbo
            .DataSource = Sociedades(cn, IncluirBlanco)
            .DisplayMember = "cpynam_0"
            .ValueMember = "cpy_0"
        End With
    End Sub
    Shared Sub SociedadesFE(ByVal cn As OracleConnection, ByVal cbo As ComboBox)
        'Selecciona Sociedades con factura electrónica
        Dim Sql As String = "SELECT cpy_0, cpynam_0, crn_0 FROM company WHERE vaccpy_0 = 'RI' ORDER BY cpy_0"
        Dim da As New OracleDataAdapter(Sql, cn)
        Dim dt As New DataTable

        da.Fill(dt)

        With cbo
            .DataSource = dt
            .DisplayMember = "cpynam_0"
            .ValueMember = "cpy_0"
        End With

    End Sub

    'PROPERTY
    Public ReadOnly Property PlantaVenta() As String
        Get
            Select Case Codigo
                Case "DNY"
                    Return "D02"
                Case "MON"
                    Return "M01"
                Case "GRU"
                    Return "G01"
                Case "LIA"
                    Return "L01"
                Case "SCH"
                    Return "S01"
                Case Else
                    Return " "
            End Select
        End Get
    End Property
    Public ReadOnly Property PlantaStock() As String
        Get
            Select Case Codigo
                Case "DNY"
                    Return "D01"
                Case "MON"
                    Return "M01"
                Case "GRU"
                    Return "G01"
                Case "LIA"
                    Return "L01"
                Case "SCH"
                    Return "S01"
                Case Else
                    Return " "
            End Select
        End Get
    End Property
    Public ReadOnly Property TieneDatos() As Boolean
        Get
            Return dt.Rows.Count > 0
        End Get
    End Property
    Public ReadOnly Property Codigo() As String
        Get
            Dim dr As DataRow
            If dt.Rows.Count = 1 Then
                dr = dt.Rows(0)
                Return dr("cpy_0").ToString
            Else
                Return Nothing

            End If
        End Get
    End Property
    Public ReadOnly Property Nombre() As String
        Get
            Dim dr As DataRow
            If dt.Rows.Count = 1 Then
                dr = dt.Rows(0)
                Return dr("cpynam_0").ToString
            Else
                Return Nothing

            End If
        End Get
    End Property
    Public ReadOnly Property Cuit() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("crn_0").ToString
        End Get
    End Property

End Class