Imports System.Data.OracleClient
Imports System.Windows.Forms

Public Class PersonaContacto
    Private cn As OracleConnection
    Private da1 As OracleDataAdapter
    Private da2 As OracleDataAdapter
    Private dt1 As DataTable
    Private dt2 As DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String

        Me.cn = cn

        Sql = "SELECT * FROM contactcrm WHERE cntnum_0 = :cntnum_0"
        da1 = New OracleDataAdapter(Sql, cn)
        da1.SelectCommand.Parameters.Add("cntnum_0", OracleType.VarChar)
        da1.InsertCommand = New OracleCommandBuilder(da1).GetInsertCommand
        da1.UpdateCommand = New OracleCommandBuilder(da1).GetUpdateCommand
        da1.DeleteCommand = New OracleCommandBuilder(da1).GetDeleteCommand

        Sql = "SELECT * FROM contact WHERE ccncrm_0 = :ccncrm_0 AND bpanum_0 = :bpanum_0 AND bpaadd_0 = :bpaadd_0"
        da2 = New OracleDataAdapter(Sql, cn)
        da2.SelectCommand.Parameters.Add("ccncrm_0", OracleType.VarChar)
        da2.SelectCommand.Parameters.Add("bpanum_0", OracleType.VarChar)
        da2.SelectCommand.Parameters.Add("bpaadd_0", OracleType.VarChar)
        da2.InsertCommand = New OracleCommandBuilder(da2).GetInsertCommand
        da2.UpdateCommand = New OracleCommandBuilder(da2).GetUpdateCommand
        da2.DeleteCommand = New OracleCommandBuilder(da2).GetDeleteCommand

    End Sub
    Public Function Abrir(ByVal CodigoContacto As String, ByVal Cliente As String, ByVal Sucursal As String) As Boolean
        If dt1 Is Nothing Then dt1 = New DataTable
        If dt2 Is Nothing Then dt2 = New DataTable

        da1.SelectCommand.Parameters("cntnum_0").Value = CodigoContacto
        da1.Fill(dt1)

        da2.SelectCommand.Parameters("ccncrm_0").Value = CodigoContacto
        da2.SelectCommand.Parameters("bpanum_0").Value = Cliente
        da2.SelectCommand.Parameters("bpaadd_0").Value = Sucursal
        da2.Fill(dt2)

        Return (dt1.Rows.Count = 1 And dt2.Rows.Count = 1)

    End Function
    Public Sub Grabar()
        da1.Update(dt1)
        da2.Update(dt2)
    End Sub

    Shared Sub LlenarComboBox(ByVal cn As OracleConnection, ByVal Cliente As String, ByVal Sucursal As String, ByVal cbo As combobox)
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim Sql As String

        Sql = "SELECT contact.*, cntlna_0 || ' ' || cntfna_0 as nombre, cntmob_0, cntlna_0, cntfna_0 "
        Sql &= "FROM contact INNER JOIN contactcrm ON (ccncrm_0 = cntnum_0) "
        Sql &= "WHERE bpanum_0 = :bpanum_0 AND bpaadd_0 = :bpaadd_0 "
        Sql &= "ORDER BY cntflg_0 DESC"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("bpanum_0", OracleType.VarChar).Value = Cliente
        da.SelectCommand.Parameters.Add("bpaadd_0", OracleType.VarChar).Value = Sucursal
        da.Fill(dt)

        With cbo
            .DataSource = dt
            .DisplayMember = "nombre"
            .ValueMember = "ccncrm_0"
        End With

    End Sub
    Public ReadOnly Property Codigo() As String
        Get
            Dim dr As DataRow

            If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                Return dr("cntnum_0").ToString
            Else
                Return ""
            End If
        End Get
    End Property
    Public ReadOnly Property NombreCompleto() As String
        Get
            Dim dr As DataRow

            If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                Return dr("cntlna_0").ToString & " " & dr("cntfna_0").ToString
            Else
                Return ""
            End If
        End Get
    End Property
    Public Property Telefono() As String
        Get
            Dim dr As DataRow

            If dt2 IsNot Nothing AndAlso dt2.Rows.Count > 0 Then
                dr = dt2.Rows(0)
                Return dr("tel_0").ToString
            Else
                Return ""
            End If
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If value.Trim = "" Then value = " "

            If dt2 IsNot Nothing AndAlso dt2.Rows.Count > 0 Then
                dr = dt2.Rows(0)
                dr.BeginEdit()
                dr("tel_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property Celular() As String
        Get
            Dim dr As DataRow

            If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                Return dr("cntmob_0").ToString
            Else
                Return ""
            End If
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If value.Trim = "" Then value = " "

            If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("cntmob_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property Mail() As String
        Get
            Dim dr As DataRow

            If dt2 IsNot Nothing AndAlso dt2.Rows.Count > 0 Then
                dr = dt2.Rows(0)
                Return dr("web_0").ToString
            Else
                Return ""
            End If
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If value.Trim = "" Then value = " "

            If dt2 IsNot Nothing AndAlso dt2.Rows.Count > 0 Then
                dr = dt2.Rows(0)
                dr.BeginEdit()
                dr("web_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property Apellido() As String
        Get
            Dim dr As DataRow

            If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                Return dr("cntlna_0").ToString
            Else
                Return ""
            End If
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If value.Trim = "" Then value = " "

            If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("cntlna_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property Nombre() As String
        Get
            Dim dr As DataRow

            If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                Return dr("cntfna_0").ToString
            Else
                Return ""
            End If
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If value.Trim = "" Then value = " "

            If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("cntfna_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property

End Class