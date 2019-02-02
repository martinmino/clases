Imports System.Data.OracleClient

Public Class Seguimiento
    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable
    Private itn As Intervencion
    Private rto As Remito

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn

        Adaptadores()

    End Sub
    Private Sub Adaptadores()
        Dim Sql As String

        Sql = "SELECT * FROM xsegto WHERE itn_0 = :nro OR rto_0 = :nro ORDER BY fecha_0, hora_0"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("nro", OracleType.VarChar)
        da.InsertCommand = New OracleCommandBuilder(da).GetInsertCommand
        da.UpdateCommand = New OracleCommandBuilder(da).GetUpdateCommand

    End Sub
    Public Function Abrir(ByVal itn As Intervencion) As Boolean
        Me.itn = itn

        da.SelectCommand.Parameters("nro").Value = itn.Numero
        dt.Clear()
        da.Fill(dt)
        Return dt.Rows.Count > 0
    End Function
    Public Function Abrir(ByVal rto As Remito) As Boolean
        Me.rto = rto

        da.SelectCommand.Parameters("nro").Value = rto.Numero
        dt.Clear()
        da.Fill(dt)
        Return dt.Rows.Count > 0
    End Function
    Public Sub MarcarTodoRecibido(ByVal Usr As String)
        Dim dr As DataRow

        For Each dr In dt.Rows
            If CInt(dr("recibido_0")) = 2 Then Continue For

            dr.BeginEdit()
            dr("recibido_0") = 2
            dr("usrrecibe_0") = Usr
            dr.EndEdit()
        Next

    End Sub
    Public Sub Grabar()
        Try
            da.Update(dt)

        Catch ex As Exception

        End Try
    End Sub
    Public Function EnviarA(ByVal Sector As String, ByVal Usuario As String, Optional ByVal AutoRecibe As Boolean = False) As Boolean
        'Salgo si el documento ya se encuentra en el sector
        If UltimoSectorDestino = Sector Then Return False
        'Salgo si el documento aun no tiene historial de seguimiento
        If Me.Count = 0 Then Return False
        'Marco el documento como recibo
        MarcarTodoRecibido("ADMIN")
        'Envio el documento al nuevo sector
        Dim dr As DataRow
        dr = dt.NewRow
        dr("fecha_0") = Date.Today
        dr("hora_0") = Now.ToString("HHmm")
        dr("usr_0") = Usuario
        dr("numlig_0") = SiguienteLinea()
        If itn IsNot Nothing Then
            dr("itn_0") = itn.Numero
            dr("rto_0") = itn.Remito
            dr("ped_0") = " "
            dr("bpcnum_0") = itn.Cliente.Codigo
            dr("bpcnam_0") = itn.Cliente.Nombre
        Else
            dr("itn_0") = " "
            dr("rto_0") = rto.Numero
            dr("ped_0") = rto.Pedido.Numero
            dr("bpcnum_0") = rto.Cliente.Codigo
            dr("bpcnam_0") = rto.Cliente.Nombre
        End If
        dr("de_0") = Me.UltimoSectorDestino
        dr("para_0") = Sector
        If AutoRecibe Then
            dr("recibido_0") = 2
            dr("usrrecibe_0") = Usuario
        End If
        dt.Rows.Add(dr)
        Grabar()

        'Actualizo el sector en el documento
        If itn IsNot Nothing Then
            itn.Sector = Sector
            itn.Grabar()
        Else
            rto.Sector = Sector
            rto.Grabar()
        End If

    End Function
    Private Function SiguienteLinea() As Integer
        Dim i As Integer = 0

        For Each dr As DataRow In dt.Rows
            If CInt(dr("numlig_0")) > i Then i = CInt(dr("numlig_0"))
        Next
        i += 1
        Return i
    End Function
    Public Function UltimaFechaEnviadoA(ByVal Sector As String) As Date
        Dim a(0) As String

        a(0) = Sector

        Return UltimaFechaEnviadoA(a)

    End Function
    Public Function UltimaFechaEnviadoA(ByVal Sector() As String) As Date
        Dim f As Date = Nothing

        For Each dr As DataRow In dt.Rows
            For i As Integer = 0 To UBound(Sector)
                If dr("para_0").ToString = Sector(i) Then
                    f = CDate(dr("fecha_0"))
                End If
            Next
        Next

        Return f
    End Function

    Public ReadOnly Property Count() As Integer
        Get
            Return dt.Rows.Count
        End Get
    End Property
    Public ReadOnly Property UltimoSectorDestino() As String
        Get
            Dim dr As DataRow

            If dt.Rows.Count = 0 Then
                Return ""
            Else
                dr = dt.Rows(dt.Rows.Count - 1)
                Return dr("para_0").ToString
            End If
        End Get
    End Property
    Public ReadOnly Property UltimoSectorOrigen() As String
        Get
            Dim dr As DataRow

            If dt.Rows.Count = 0 Then
                Return ""
            Else
                dr = dt.Rows(dt.Rows.Count - 1)
                Return dr("de_0").ToString
            End If
        End Get
    End Property

End Class