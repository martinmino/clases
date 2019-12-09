Imports System.Data.OracleClient

Public Class Camioneta
    Private da As OracleDataAdapter
    Private dt As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Dim sql As String

        sql = "SELECT * "
        sql &= "FROM zunitrans "
        sql &= "WHERE bptnum_0 = :bptnum AND patnum_0 = :patnum"

        da = New OracleDataAdapter(sql, cn)
        With da.SelectCommand.Parameters
            .Add("bptnum", OracleType.VarChar)
            .Add("patnum", OracleType.VarChar)
        End With

    End Sub
    Public Function Abrir(ByVal Empresa As String, ByVal Patente As String) As Boolean
        da.SelectCommand.Parameters("bptnum").Value = Empresa
        da.SelectCommand.Parameters("patnum").Value = Patente

        Try
            dt.Clear()
            da.Fill(dt)

        Catch ex As Exception
            Return False

        End Try

        Return True

    End Function
    Public ReadOnly Property Codigo() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("bptnum_0").ToString
        End Get
    End Property
    Public Property ChoferInterno() As Boolean
        Get
            Dim dr As DataRow
            Dim f As Boolean = False

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                f = CBool(IIf(CInt(dr("interno_0")) <> 2, False, True))
            End If

            Return f
        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("interno_0") = IIf(value, 2, 1)
                dr.EndEdit()
            End If
        End Set
    End Property
    Public ReadOnly Property Patente() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("patnum_0").ToString
        End Get
    End Property
    Public ReadOnly Property Sector() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CInt(dr("xsector_0"))
        End Get
    End Property
    Public ReadOnly Property Acomp() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CInt(dr("acomp_0"))
        End Get
    End Property

End Class