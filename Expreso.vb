Imports System.Data.OracleClient
Imports System.Windows.Forms

Public Class Expreso
    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String

        Me.cn = cn

        Sql = "SELECT * FROM xtran WHERE tranum_0 = :tranum"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("tranum", OracleType.VarChar)

    End Sub
    Public Sub Grabar()
        Try
            da.Update(dt)
        Catch ex As Exception
        End Try
    End Sub
    Public Function Abrir(ByVal Codigo As String) As Boolean
        da.SelectCommand.Parameters("tranum").Value = Codigo
        dt.Clear()
        da.Fill(dt)
        Return dt.Rows.Count > 0
    End Function
    Public Function Existe(ByVal Codigo As String) As Boolean
        Return Abrir(Codigo)
    End Function
    Public Sub Enlazar(ByVal cbo As ComboBox, Optional ByVal Blanco As Boolean = False)
        Dim dt As DataTable
        Dim dr As DataRow

        dt = ObtenerTodos()

        If Blanco Then
            dr = dt.NewRow
            dr("tranum_0") = " "
            dr("nombre_0") = " "
            dr("domicilio_0") = " "
            dr("ciudad_0") = " "
            dt.Rows.InsertAt(dr, 0)
        End If

        With cbo
            .ValueMember = "tranum_0"
            .DisplayMember = "nombre_0"
            .DataSource = dt
        End With

    End Sub
    Public Function ObtenerTodos() As DataTable
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable

        Sql = "SELECT tranum_0, nombre_0, domicilio_0, ciudad_0 FROM xtran ORDER BY nombre_0"
        da = New OracleDataAdapter(Sql, cn)
        dt.Clear()
        da.Fill(dt)
        Return dt
    End Function
    Public Property Codigo() As String
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return dr("tranum_0").ToString
            Else
                Return ""
            End If
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("tranum_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property Nombre() As String
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return dr("nombre_0").ToString
            Else
                Return ""
            End If
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("nombre_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property

End Class
