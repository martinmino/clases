Imports System.Data.OracleClient
Imports System.Windows.Forms

Public Class TablaVaria
    Private cn As OracleConnection
    Private da1 As OracleDataAdapter
    Private da2 As OracleDataAdapter
    Private dt1 As New DataTable
    Private dt2 As New DataTable

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String

        Me.cn = cn

        Sql = "SELECT * FROM atabdiv WHERE numtab_0 = :numtab"
        da1 = New OracleDataAdapter(Sql, cn)
        da1.SelectCommand.Parameters.Add("numtab", OracleType.Number)

        Sql = "SELECT ate.*, ident2_0 || ' - ' || texte_0 AS Combi, texte_0 FROM atextra ate WHERE codfic_0 = 'ATABDIV' AND zone_0 = 'LNGDES' AND langue_0 = 'SPA' AND ident1_0 = :numtab"
        da2 = New OracleDataAdapter(Sql, cn)
        da2.SelectCommand.Parameters.Add("numtab", OracleType.VarChar)

    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal NroTabla As Integer, Optional ByVal Blanco As Boolean = False)
        Me.new(cn)
        AbrirTabla(NroTabla, Blanco)

    End Sub
    Public Sub AbrirTabla(ByVal NroTabla As Integer, Optional ByVal Blanco As Boolean = False)
        Dim dr As DataRow

        da1.SelectCommand.Parameters("numtab").Value = NroTabla
        da2.SelectCommand.Parameters("numtab").Value = NroTabla.ToString

        Try
            dt1.Clear()
            dt2.Clear()
            da1.Fill(dt1)
            da2.Fill(dt2)

            If Blanco Then
                dr = dt2.NewRow
                dr("codfic_0") = "ATABDIV"
                dr("zone_0") = "LNGDES"
                dr("langue_0") = "SPA"
                dr("ident1_0") = NroTabla
                dr("ident2_0") = " "
                dr("texte_0") = " "
                dr("combi") = " "
                dt2.Rows.InsertAt(dr, 0)
                dt2.AcceptChanges()
            End If

        Catch ex As Exception
        End Try

    End Sub
    Public Sub EnlazarComboBox(ByVal cbo As ComboBox)
        With cbo
            .DataSource = dt2
            .DisplayMember = "combi"
            .ValueMember = "ident2_0"
        End With
    End Sub
    Public Sub EnlazarComboBox(ByVal cbo As DataGridViewComboBoxColumn)
        With cbo
            .DataSource = dt2
            .DisplayMember = "texte_0"
            .ValueMember = "ident2_0"
        End With
    End Sub
    Public Sub EnlazarListBox(ByVal cbo As ListBox)
        With cbo
            .DataSource = dt2
            .DisplayMember = "combi"
            .ValueMember = "ident2_0"
        End With
    End Sub

    Public ReadOnly Property Aux1(ByVal Code As String) As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            For Each dr In dt1.Rows
                If dr("code_0").ToString = Code Then
                    txt = dr("a1_0").ToString
                End If
            Next

            Return txt
        End Get
    End Property
    Public ReadOnly Property Texto(ByVal Code As String) As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            For Each dr In dt2.Rows
                If dr("ident2_0").ToString = Code Then
                    txt = dr("texte_0").ToString
                End If
            Next

            Return txt
        End Get
    End Property

    Public ReadOnly Property Tabla1() As DataTable
        Get
            Return dt1
        End Get
    End Property
    Public ReadOnly Property Tabla2() As DataTable
        Get
            Return dt2
        End Get
    End Property

End Class