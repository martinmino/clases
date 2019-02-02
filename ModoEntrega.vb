Imports System.Data.OracleClient
Imports System.Windows.Forms

Public Class ModoEntrega
    Private dt As New DataTable

    Public Sub New(ByVal cn As OracleConnection, Optional ByVal Blanco As Boolean = False)
        Dim da As OracleDataAdapter
        Dim dr As DataRow
        Dim sql As String
        Dim spl() As String

        sql = "SELECT mdl_0, landessho_0 FROM tabmodeliv ORDER BY mdl_0"
        da = New OracleDataAdapter(sql, cn)
        da.Fill(dt)
        da.Dispose()

        If Blanco Then
            dr = dt.NewRow
            dr(0) = "0"
            dr(1) = " "
            dt.Rows.InsertAt(dr, 0)
        End If

        For Each dr In dt.Rows
            If CInt(dr(0)) > 2 Then
                dr.Delete()
                Continue For
            End If

            If dr("landessho_0").ToString.IndexOf("~") > -1 Then

                spl = Split(dr("landessho_0").ToString, "~")

                dr.BeginEdit()
                dr("landessho_0") = spl(1)
                dr.EndEdit()
            End If
        Next

        dt.AcceptChanges()

    End Sub
    Public Sub ModosEntrega(ByVal cbo As ComboBox)
        With cbo
            cbo.DataSource = dt
            cbo.DisplayMember = "landessho_0"
            cbo.ValueMember = "mdl_0"
        End With
    End Sub
    Public ReadOnly Property TipoEntrega(ByVal Code As String) As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            For Each dr In dt.Rows
                If dr(0).ToString = Code Then
                    txt = dr(1).ToString
                    Exit For
                End If
            Next

            Return txt

        End Get
    End Property

End Class
