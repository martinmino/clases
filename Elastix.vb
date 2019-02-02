Imports System.Data.OracleClient

Public Class Elastix
    Private cn As OracleConnection

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
    End Sub
    Public Sub Registrar(ByVal Cliente As String, ByVal Telefono As String)
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt1 As New DataTable
        Dim dt2 As New DataTable
        Dim flg As Boolean = False

        'Consulto si existe el cliente
        Sql = "SELECT bpc.bpcnum_0, rep.tel_0 AS rep, aus.tel_0 AS aus "
        Sql &= "FROM bpcustomer bpc INNER JOIN "
        Sql &= "     salesrep rep ON (bpc.rep_0 = rep.repnum_0) INNER JOIN "
        Sql &= "     autilis aus ON (rep.xanalis_0 = aus.usr_0) "
        Sql &= "WHERE (bpcnum_0 = :cliente OR docnum_0 = :cliente) AND "
        Sql &= "      bpcsta_0 = 2 AND rep_0 <> '07'"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("cliente", OracleType.VarChar).Value = Cliente
        da.Fill(dt1)

        'Consulta para insertar datos
        Sql = "SELECT * FROM elastix WHERE bpcnum_0 = :cliente AND tel_0 = :telefono"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("cliente", OracleType.VarChar)
        da.SelectCommand.Parameters.Add("telefono", OracleType.VarChar)
        da.InsertCommand = New OracleCommandBuilder(da).GetInsertCommand

        For Each dr1 As DataRow In dt1.Rows

            'Busco si existe el cliente/telefono
            da.SelectCommand.Parameters("cliente").Value = dr1("bpcnum_0").ToString
            da.SelectCommand.Parameters("telefono").Value = Telefono
            dt2.Clear()
            da.Fill(dt2)

            If dt2.Rows.Count = 0 Then
                Dim dr2 As DataRow
                dr2 = dt2.NewRow
                dr2(0) = Cliente
                dr2(1) = Telefono
                dt2.Rows.Add(dr2)
                da.Update(dt2)
            End If

        Next

    End Sub
    Public Function Borrar(ByVal Tel As String) As Integer
        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable
        Dim i As Integer

        'Busco numero de telefono
        Sql = "SELECT * FROM elastix WHERE tel_0 = :telefono"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("telefono", OracleType.VarChar).Value = Tel
        da.DeleteCommand = New OracleCommandBuilder(da).GetDeleteCommand

        da.Fill(dt)

        i = dt.Rows.Count

        For Each dr As DataRow In dt.Rows
            dr.Delete()
        Next

        If i > 0 Then da.Update(dt)

        Return i

    End Function
    Public Function ObtenerVendedorPorCliente(ByVal Valor As String) As String
        Dim da As OracleDataAdapter
        Dim Sql As String
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim Interno As String = "0"

        'Busco si está registrado el telefono y obtengo el interno del vendedor.
        Sql = "SELECT DISTINCT rep.tel_0 "
        Sql &= "FROM bpcustomer bpc INNER JOIN "
        Sql &= "	 salesrep  rep ON (bpc.rep_0 = rep.repnum_0) "
        Sql &= "WHERE bpc.bpcnum_0 = :valor OR "
        Sql &= "      bpc.docnum_0 = :valor"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("valor", OracleType.VarChar).Value = Valor
        da.Fill(dt)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            Interno = dr(0).ToString.Trim
        End If

        If Interno = "" Then Interno = "0"

        Return Interno.Trim

    End Function
    Public Function ObtenerAsistentePorCliente(ByVal Valor As String) As String
        Dim da As OracleDataAdapter
        Dim Sql As String
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim Interno As String = "0"

        'Busco si está registrado el telefono y obtengo el interno del vendedor.
        Sql = "SELECT DISTINCT aus.tel_0 "
        Sql &= "FROM bpcustomer bpc INNER JOIN "
        Sql &= "	 salesrep rep ON (bpc.rep_0 = rep.repnum_0) INNER JOIN "
        Sql &= "	 autilis aus ON (rep.xanalis_0 = aus.usr_0) "
        Sql &= "WHERE bpc.bpcnum_0 = :valor OR "
        Sql &= "      bpc.docnum_0 = :valor"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("valor", OracleType.VarChar).Value = Valor
        da.Fill(dt)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            Interno = dr(0).ToString
        End If

        If Interno.Trim = "" Then Interno = "0"

        Return Interno.Trim

    End Function
    Public Function ObtenerVendedorPorTelefono(ByVal Valor As String) As String
        Dim da As OracleDataAdapter
        Dim Sql As String
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim Interno As String = ""
        Dim Interno07 As String = ""

        'Busco si está registrado el telefono y obtengo el interno del vendedor
        Sql = "SELECT DISTINCT bpc.rep_0, rep.tel_0 "
        Sql &= "FROM elastix ela INNER JOIN "
        Sql &= "	 bpcustomer bpc ON (ela.bpcnum_0 = bpc.bpcnum_0) INNER JOIN "
        Sql &= "	 salesrep  rep ON (bpc.rep_0 = rep.repnum_0) "
        Sql &= "WHERE ela.tel_0 = :telefono AND "
        Sql &= "      bpc.bpcsta_0 = 2 "

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("telefono", OracleType.VarChar).Value = Valor
        da.Fill(dt)

        For Each dr In dt.Rows
            If dr("rep_0").ToString = "07" Then
                Interno07 = dr("tel_0").ToString
            Else
                Interno = dr("tel_0").ToString
            End If
        Next

        If Interno = "" Then Interno = Interno07
        If Interno = "" Then Interno = "0"

        Return Interno

    End Function
    Public Function ObtenerAsistentePorTelefono(ByVal Valor As String) As String
        Dim da As OracleDataAdapter
        Dim Sql As String
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim Interno As String = ""
        Dim Interno07 As String = ""

        'Busco si está registrado el telefono y obtengo el interno del vendedor
        Sql = "SELECT DISTINCT bpc.rep_0, aus.tel_0 "
        Sql &= "FROM elastix ela INNER JOIN "
        Sql &= "	 bpcustomer bpc ON (ela.bpcnum_0 = bpc.bpcnum_0) INNER JOIN "
        Sql &= "	 salesrep   rep ON (bpc.rep_0 = rep.repnum_0) INNER JOIN "
        Sql &= "	 autilis    aus ON (rep.xanalis_0 = aus.usr_0) "
        Sql &= "WHERE ela.tel_0 = :telefono AND "
        Sql &= "      bpc.bpcsta_0 = 2 "

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("telefono", OracleType.VarChar).Value = Valor
        da.Fill(dt)

        For Each dr In dt.Rows
            If dr("rep_0").ToString = "07" Then
                Interno07 = dr("tel_0").ToString
            Else
                Interno = dr("tel_0").ToString
            End If
        Next

        If Interno = "" Then Interno = Interno07
        If Interno = "" Then Interno = "0"

        Return Interno

    End Function

End Class