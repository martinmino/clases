Imports System.Data.OracleClient

Public Class ReporteRecargas
    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable

    Private Equipos(1) As Integer
    Private Sustitutos(1) As Integer
    Private Rechazos(1) As Integer

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
    End Sub
    Public Sub BuscarRetirados(ByVal Fecha As Date)
        Dim Sql As String

        Sql = "SELECT * FROM xretiros WHERE dat_0 = :dat"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("dat", OracleType.DateTime).Value = Fecha
        dt.Clear()
        da.Fill(dt)

        If dt.Rows.Count = 1 Then
            Dim dr As DataRow
            dr = dt.Rows(0)
            Equipos(0) = CInt(dr("extintores_0"))
            Equipos(1) = CInt(dr("mangueras_0"))
        Else
            Equipos(0) = 0
            Equipos(1) = 0
        End If
    End Sub
    Public Sub BuscarRecepcionados(ByVal Fecha As Date)
        Dim Sql As String

        Sql = "SELECT xsg.itn_0, itm.tsicod_3, mac.macitntyp_0, mac.bpcnum_0 "
        Sql &= "FROM xsegto2 xsg inner join "
        Sql &= "	 interven itn on (xsg.itn_0 = itn.num_0) inner join "
        Sql &= "	 sremac srm on (xsg.itn_0 = srm.yitnnum_0) inner join "
        Sql &= "	 machines mac on (srm.macnum_0 = mac.macnum_0) inner join "
        Sql &= "	 itmmaster itm on (mac.macpdtcod_0 = itm.itmref_0) "
        Sql &= "WHERE srm.credat_0 = :dat and "
        Sql &= "      srm.creusr_0 = 'RECEP'"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("dat", OracleType.DateTime).Value = Fecha

        dt.Clear()
        da.Fill(dt)

        SumarCantidades()
    End Sub   
    Public Sub BuscarAPrefacturacion(ByVal Fecha As Date)
        Buscar(Fecha, 2)
    End Sub
    Public Sub BuscarALogistica(ByVal Fecha As Date)
        Buscar(Fecha, 3)
    End Sub
    Public Sub BuscarEntregados(ByVal Fecha As Date)
        Buscar(Fecha, 4)
    End Sub
    Private Sub Buscar(ByVal Fecha As Date, ByVal Campo As Byte)
        Dim Sql As String

        Sql = "SELECT xsg.itn_0, itm.tsicod_3, mac.macitntyp_0, mac.bpcnum_0 "
        Sql &= "FROM xsegto2 xsg inner join "
        Sql &= "	 interven itn on (xsg.itn_0 = itn.num_0) inner join "
        Sql &= "	 sremac srm on (xsg.itn_0 = srm.yitnnum_0) inner join "
        Sql &= "	 machines mac on (srm.macnum_0 = mac.macnum_0) inner join "
        Sql &= "	 itmmaster itm on (mac.macpdtcod_0 = itm.itmref_0) "
        Sql &= "WHERE xsg.dat_? = :dat and "
        Sql &= "	  srm.creusr_0 = 'RECEP'"

        'Pongo el indice del campo por el que se va a consultar
        Sql = Sql.Replace("?", Campo.ToString)

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("dat", OracleType.DateTime).Value = Fecha

        dt.Clear()
        da.Fill(dt)

        SumarCantidades()
    End Sub
    Private Sub SumarCantidades()

        'Pongo variables a cero
        For i = 0 To 1
            Equipos(i) = 0
            Sustitutos(i) = 0
            Rechazos(i) = 0
        Next

        For Each dr As DataRow In dt.Rows
            If dr("bpcnum_0").ToString = "402000" Then

                Select Case dr("tsicod_3").ToString
                    Case "301" 'Extintores
                        Sustitutos(0) += 1

                    Case "302" 'Mangueras
                        Sustitutos(1) += 1

                End Select

            Else

                Select Case dr("tsicod_3").ToString
                    Case "301" 'Extintores
                        If CInt(dr("macitntyp_0")) = 1 Then
                            Equipos(0) += 1
                        Else
                            Rechazos(0) += 1
                        End If


                    Case "302" 'Mangueras
                        If CInt(dr("macitntyp_0")) = 1 Then
                            Equipos(1) += 1
                        Else
                            Rechazos(1) += 1
                        End If

                End Select

            End If
        Next

    End Sub

    Public ReadOnly Property Extintores() As Integer
        Get
            Return Equipos(0)
        End Get
    End Property
    Public ReadOnly Property Mangueras() As Integer
        Get
            Return Equipos(1)
        End Get
    End Property
    Public ReadOnly Property ExtintoresSustitutos() As Integer
        Get
            Return Sustitutos(0)
        End Get
    End Property
    Public ReadOnly Property ManguerasSustitutas() As Integer
        Get
            Return Sustitutos(1)
        End Get
    End Property
    Public ReadOnly Property ExtintoresRechazados() As Integer
        Get
            Return Rechazos(0)
        End Get
    End Property
    Public ReadOnly Property ManguerasRechazadas() As Integer
        Get
            Return Rechazos(1)
        End Get
    End Property
    Public ReadOnly Property TotalExtintores() As Integer
        Get
            Return Equipos(0) + Rechazos(0) + Sustitutos(0)
        End Get
    End Property
    Public ReadOnly Property TotalMangueras() As Integer
        Get
            Return Equipos(1) + Rechazos(1) + Sustitutos(1)
        End Get
    End Property

End Class