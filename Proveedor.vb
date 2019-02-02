Imports System.Data.OracleClient

Public Class Proveedor
    Inherits Tercero

    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        MyBase.New(cn)
        Me.cn = cn

        Dim Sql As String

        Sql = "SELECT * FROM bpsupplier WHERE bpsnum_0 = :bpsnum_0"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("bpsnum_0", OracleType.VarChar)

    End Sub
    Public Overloads Function Abrir(ByVal Codigo As String) As Boolean
        da.SelectCommand.Parameters("bpsnum_0").Value = Codigo
        dt.Clear()
        da.Fill(dt)

        MyBase.Abrir(Codigo)
        Return dt.Rows.Count > 0
    End Function
    Public Overloads Sub Grabar()
        MyBase.Grabar()
        da.Update(dt)
    End Sub

    Public Property MailAvisoPago() As String
        Get
            Dim dr As DataRow

            If dt IsNot Nothing And dt.Rows.Count = 1 Then
                dr = dt.Rows(0)
                Return dr("xmailaviso_0").ToString.Trim

            Else
                Return ""

            End If
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            value = value.Trim
            If value = "" Then value = " "

            If dt IsNot Nothing And dt.Rows.Count = 1 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("xmailaviso_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property

End Class
