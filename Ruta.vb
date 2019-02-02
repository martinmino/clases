Imports System.Data.OracleClient

Public Class Ruta
    Private cn As OracleConnection
    Private dt1 As DataTable 'Rutac
    Private dt2 As DataTable 'Rutad
    Private da1 As OracleDataAdapter 'Rutac
    Private da2 As OracleDataAdapter 'Rutad
    Private Motivos As New TablaVaria(cn)

    Private dv As DataView

    Public Sub New(ByVal cn As OracleConnection)
        Dim Sql As String

        Me.cn = cn

        Sql = "SELECT * FROM xrutac where ruta_0 = :ruta"
        da1 = New OracleDataAdapter(Sql, cn)
        da1.SelectCommand.Parameters.Add("ruta", OracleType.Number)
        da1.InsertCommand = New OracleCommandBuilder(da1).GetInsertCommand
        da1.UpdateCommand = New OracleCommandBuilder(da1).GetUpdateCommand
        da1.DeleteCommand = New OracleCommandBuilder(da1).GetDeleteCommand

        Sql = "SELECT * FROM xrutad where ruta_0 = :ruta"
        da2 = New OracleDataAdapter(Sql, cn)
        da2.SelectCommand.Parameters.Add("ruta", OracleType.Number)
        da2.InsertCommand = New OracleCommandBuilder(da2).GetInsertCommand
        da2.UpdateCommand = New OracleCommandBuilder(da2).GetUpdateCommand
        da2.DeleteCommand = New OracleCommandBuilder(da2).GetDeleteCommand


        Motivos.AbrirTabla(5000)

    End Sub

    Public Sub Abrir(ByVal Numero As Integer)
        If dt1 Is Nothing Then dt1 = New DataTable
        If dt2 Is Nothing Then dt2 = New DataTable
        dt1.Clear()
        dt2.Clear()

        da1.SelectCommand.Parameters("ruta").Value = Numero
        da2.SelectCommand.Parameters("ruta").Value = Numero

        Try
            da1.Fill(dt1)
            da2.Fill(dt2)

        Catch ex As Exception

        End Try

    End Sub
    Public Function AbrirUltimaRuta(ByVal NumeroDocumento As String, Optional ByVal ExcluirRuta As Integer = 0) As Boolean
        Dim dt As New DataTable
        Dim da As OracleDataAdapter
        Dim sql As String = " "

        sql = "SELECT * "
        sql &= "FROM xrutad "
        sql &= "WHERE vcrnum_0 = :num "
        If ExcluirRuta > 0 Then
            sql &= " and ruta_0 = " & ExcluirRuta.ToString
        End If
        sql &= " ORDER BY ruta_0"

        da = New OracleDataAdapter(sql, cn)
        da.SelectCommand.Parameters.Add("num", OracleType.VarChar).Value = NumeroDocumento
        da.Fill(dt)

        If dt.Rows.Count > 0 Then
            Dim dr As DataRow

            dr = dt.Rows(0)

            Abrir(CInt(dr("ruta_0")))

            Return True

        End If

        Return False

    End Function
    'Public Function TablaMotivos(ByVal motivo As String) As String
    '    Dim da3 As New OracleDataAdapter("SELECT texte_0 FROM atabdiv INNER JOIN atextra ON (numtab_0 = ident1_0) AND (code_0 = ident2_0) WHERE numtab_0 = 5000 AND codfic_0 = 'ATABDIV' AND zone_0 = 'LNGDES' and code_0 = :codigo ORDER BY code_0", cn)
    '    Dim dt3 As New DataTable
    '    da3.SelectCommand.Parameters.Add("codigo", OracleType.VarChar).Value = motivo
    '    da3.Fill(dt3)
    '    da3.Dispose()

    '    Return dt3.Rows(0).Item("texte_0").ToString

    'End Function
    Public Function Observacion(ByVal NumeroDocumento As String) As String
        Dim txt As String = ""
        Dim dr As DataRow

        For Each dr In dt2.Rows
            If dr("vcrnum_0").ToString = NumeroDocumento Then
                txt = dr("obs_0").ToString
            End If
        Next

        Return txt

    End Function
    Public Function Sube(ByVal NumeroDocumento As String) As Integer
        Dim dr As DataRow
        Dim i As Integer = 1

        For Each dr In dt2.Rows
            If dr("vcrnum_0").ToString = NumeroDocumento Then
                i = CInt(dr("sube_0"))
            End If
        Next

        Return i

    End Function
    Public Function MotivoRebote(ByVal NumeroDocumento As String) As String
        Dim dr As DataRow
        Dim i As Integer = 1

        For Each dr In dt2.Rows
            If dr("vcrnum_0").ToString = NumeroDocumento Then
                i = CInt(dr("noconform_0"))
            End If
        Next

        Return Motivos.Texto(i.ToString)

    End Function
    Public ReadOnly Property Camioneta() As Camioneta
        Get
            If dt1 Is Nothing OrElse dt1.Rows.Count = 0 Then Return Nothing
            Dim c As New Camioneta(cn)
            Dim dr As DataRow = dt1.Rows(0)
            Dim Transporte As String = dr("transporte_0").ToString
            Dim Patente As String = dr("patente_0").ToString

            c.Abrir(Transporte, Patente)

            Return c

        End Get
    End Property
    Public ReadOnly Property EstuvoMicrocentro() As Integer
        Get
            Dim dr As DataRow = dt1.Rows(0)
            Return CInt(dr("micro_0"))
        End Get
    End Property
    Public Property Fecha() As Date
        Get
            Dim dr As DataRow
            Dim f As Date = #12/31/1599#

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                f = CDate(dr("fecha_0"))
            End If

            Return f

        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                dr.BeginEdit()
                dr("fecha_0") = value
                dr.EndEdit()

            End If

        End Set
    End Property
    Public ReadOnly Property Numero() As Integer
        Get
            Dim dr As DataRow
            Dim i As Integer = 0

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                i = CInt(dr("ruta_0"))
            End If

            Return i
        End Get
    End Property
    Public ReadOnly Property UsuarioCreacion() As String
        Get
            Dim dr As DataRow
            Dim s As String = ""

            If dt1.Rows.Count > 0 Then
                dr = dt1.Rows(0)
                s = dr("creusr_0").ToString
            End If

            Return s
        End Get
    End Property

End Class
