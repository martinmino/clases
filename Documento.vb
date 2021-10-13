Imports System.Data.OracleClient

Public Class Documento

    Private cn As OracleConnection

    Private dtSolicitud As New DataTable
    Private dtIntervencion As New DataTable
    Private dtRetiros As New DataTable
    Private dtConsumos As New DataTable
    Private dtRemito As New DataTable
    Private dtRemitod As New DataTable
    Private dtRutad As New DataTable

    Private daSolicitud As OracleDataAdapter
    Private daIntervencion As OracleDataAdapter
    Private daRetiros As OracleDataAdapter
    Private daConsumos As OracleDataAdapter
    Private daRemito As OracleDataAdapter
    Private daRemitod As OracleDataAdapter
    Private daRutad As OracleDataAdapter

    'Variable de propiedades
    Private l_Equipos As Integer = 0
    Private l_Mangas As Integer = 0
    Private l_PrestamosExt As Integer = 0
    Private l_PrestamosMan As Integer = 0
    Private l_Instalaciones As Integer = 0
    Private l_RechazosExt As Integer = 0
    Private l_RechazosMan As Integer = 0
    Private l_Varios As Boolean = False
    Private l_Peso As Double = 0
    Private l_Peso2 As Double = 0 'Peso para unigis sin prestamos
    Private l_Hora As String = " "
    Private l_TieneInstalacion As Boolean = False
    Private l_EsTarea As Boolean = False
    Private l_Serie As String = " "
    Private l_NroDocumento As String = ""

    'Eventos
    Public Event ErrorDocumento(ByVal sender As Object, ByVal e As ErrDocumentoEvenArgs)

    Public Sub New(ByRef Conexion As OracleConnection)
        cn = Conexion

        daSolicitud = New OracleDataAdapter("SELECT * FROM serrequest WHERE srenum_0 = :p1", cn)
        daSolicitud.SelectCommand.Parameters.Add("p1", OracleType.VarChar)

        daIntervencion = New OracleDataAdapter("SELECT * FROM interven WHERE num_0 = :p1", cn)
        daIntervencion.SelectCommand.Parameters.Add("p1", OracleType.VarChar)

        daConsumos = New OracleDataAdapter("SELECT itmref_0, hdtqty_0, cfglin_0, tclcod_0, hdtqty_0*itmwei_0 AS peso FROM hdktask INNER JOIN itmmaster ON (hdtitm_0 = itmref_0) WHERE itnnum_0 = :p1", cn)
        daConsumos.SelectCommand.Parameters.Add("p1", OracleType.VarChar)

        daRetiros = New OracleDataAdapter("SELECT yit.itmref_0, tqty_0, typlig_0, cfglin_0, tclcod_0, tqty_0*itmwei_0 AS peso FROM yitndet yit INNER JOIN itmmaster itm ON (yit.itmref_0 = itm.itmref_0) WHERE num_0 = :p1", cn)
        daRetiros.SelectCommand.Parameters.Add("p1", OracleType.VarChar)

        daRemito = New OracleDataAdapter("SELECT * FROM sdelivery WHERE sdhnum_0 = :p1", cn)
        daRemito.SelectCommand.Parameters.Add("p1", OracleType.VarChar)

        '*** Al implementar cambiar WHERE zhrvta_0 = 2 por zhrvta_0 = 1 ***
        daRemitod = New OracleDataAdapter("SELECT sde.itmref_0, qty_0, cfglin_0, tclcod_0, qty_0*itmwei_0 AS peso FROM sdeliveryd sde INNER JOIN itmmaster itm ON (sde.itmref_0 = itm.itmref_0) WHERE zhrvta_0 = 2 AND sdhnum_0 = :p1", cn)
        daRemitod.SelectCommand.Parameters.Add("p1", OracleType.VarChar)

        daRutad = New OracleDataAdapter("SELECT * FROM xrutad WHERE vcrnum_0 = :p1 ORDER BY ruta_0 DESC", cn)
        daRutad.SelectCommand.Parameters.Add("p1", OracleType.VarChar)

    End Sub

    Private Function NombreCliente(ByVal Codigo As String) As String
        Dim dr As OracleDataReader
        Dim cm As New OracleCommand("SELECT * FROM bpcustomer WHERE bpcnum_0 = :p1", cn)

        cm.Parameters.Add("p1", OracleType.VarChar)
        cm.Parameters(0).Value = Codigo

        dr = cm.ExecuteReader

        If dr.Read Then
            NombreCliente = dr("bpcnam_0").ToString
        Else
            NombreCliente = ""
        End If

        dr.Close()
        dr = Nothing

        cm.Dispose() : cm = Nothing

    End Function
    Private Function AnalizarIntervencionRetiro() As Boolean
        Dim dr As DataRow

        'Intervencion sin remito. Se trata de un retiro de recarga (RET)
        For Each dr In dtRetiros.Rows
            'Recorro todos los codigos cargados en la solapa Retiro de la intervención
            Select Case CType(dr("typlig_0"), Integer)
                Case 1  'Solapa Retiro

                    Select Case dr("cfglin_0").ToString
                        Case "451", "459"  'Recarga equipos
                            l_Equipos += CType(dr("tqty_0"), Integer)
                            l_Peso += CType(dr("peso"), Double)
                            l_Peso2 += CType(dr("peso"), Double)

                        Case "505"  'Mangas para PH
                            l_Mangas += CType(dr("tqty_0"), Integer)
                            l_Peso += CType(dr("peso"), Double)
                            l_Peso2 += CType(dr("peso"), Double)

                        Case "504"  'Mantenimiento hidrantes
                            l_EsTarea = True
                            l_Equipos += CType(dr("tqty_0"), Integer)

                        Case "551"  'Mantenimiento sistemas fijos
                            l_EsTarea = True
                            l_Equipos += CType(dr("tqty_0"), Integer)

                        Case "553"  'Relevamientos
                            l_EsTarea = True

                            If dr("itmref_0").ToString = "553010" Then
                                l_Mangas += CType(dr("tqty_0"), Integer)

                            Else
                                l_Equipos += CType(dr("tqty_0"), Integer)

                            End If

                        Case "652", "659"   'Controles y visitas / Otros
                            l_EsTarea = True
                            l_Equipos += CType(dr("tqty_0"), Integer)

                        Case Else
                            Dim txt As String = "El artículo {0} no se permite en la sección retiros."
                            txt = String.Format(txt, dr("itmref_0").ToString)

                            RaiseEvent ErrorDocumento(Me, New ErrDocumentoEvenArgs(txt))
                            Return False

                    End Select

                Case 2  'Solapa préstamos
                    If dr("tclcod_0").ToString = "60" Then

                        If dr("itmref_0").ToString.StartsWith("60100") Then
                            l_PrestamosExt += CType(dr("tqty_0"), Integer)
                            l_Peso += CType(dr("peso"), Double)

                        ElseIf dr("itmref_0").ToString.StartsWith("60700") Then
                            l_PrestamosMan += CType(dr("tqty_0"), Integer)
                            l_Peso += CType(dr("peso"), Double)

                        End If

                    Else
                        Dim txt As String = "El artículo {0} no se puede cargar en la sección de préstamos."
                        txt = String.Format(txt, dr("itmref_0").ToString)

                        RaiseEvent ErrorDocumento(Me, New ErrDocumentoEvenArgs(txt))
                        Return False

                    End If

                Case 3  'Solapa Instalaciones
                    Dim txt As String = "Intervención mal cargada{0}{0}No se permiten intervenciones con ítems en la sección Instalaciones.{0}Ver con Administración de Ventas."
                    txt = String.Format(txt, vbCrLf)

                    RaiseEvent ErrorDocumento(Me, New ErrDocumentoEvenArgs(txt))
                    Return False

            End Select

        Next

        Return True

    End Function
    Private Function AnalizarIntervencionEntrega() As Boolean
        Dim dr As DataRow
        Dim txt As String = ""

        'Recorro todos los consumos cargados en el remito
        For Each dr In dtConsumos.Rows
            l_Peso += CType(dr("peso"), Double)

            If dr("cfglin_0").ToString = "453" Then 'Rechazo Extintores
                l_RechazosExt += CType(dr("hdtqty_0"), Integer)
                l_Peso2 += CType(dr("peso"), Double)

            ElseIf dr("cfglin_0").ToString = "503" Then 'Rechazo Mangueras
                l_RechazosMan += CType(dr("hdtqty_0"), Integer)
                l_Peso2 += CType(dr("peso"), Double)

            ElseIf dr("cfglin_0").ToString = "451" Then 'Extintores
                l_Equipos += CType(dr("hdtqty_0"), Integer)
                l_Peso2 += CType(dr("peso"), Double)

            ElseIf dr("cfglin_0").ToString = "505" Then 'Mangueras
                l_Mangas += CType(dr("hdtqty_0"), Integer)
                l_Peso2 += CType(dr("peso"), Double)

            ElseIf dr("tclcod_0").ToString = "60" Then

                If dr("itmref_0").ToString.StartsWith("60100") Then
                    l_PrestamosExt += CType(dr("hdtqty_0"), Integer)

                ElseIf dr("itmref_0").ToString.StartsWith("60700") Then
                    l_PrestamosMan += CType(dr("hdtqty_0"), Integer)

                End If

            End If

        Next

        '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        'PROGRAMAR AQUI LA COMPARACION DE CANTIDAD DE PRESTAMOS DE REMITO CONTRA PARQUE
        '¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡

        'Consulto parque de prestamos dejado al cliente (EXTINTORES)
        Dim ddr As OracleDataReader
        Dim cm As New OracleCommand("SELECT * FROM machines WHERE ynrocil_0 = :ynrocil_0 AND macpdtcod_0 = :macpdtcod_0", cn)
        cm.Parameters.Add("ynrocil_0", OracleType.VarChar).Value = Me.Documento


        Try
            'Consulto el parque de extintores de prestamos
            cm.Parameters.Add("macpdtcod_0", OracleType.VarChar).Value = ARTICULO_PRESTAMO_EXT
            ddr = cm.ExecuteReader

            If ddr.Read Then

                If CInt(ddr("macqty_0")) <> l_PrestamosExt Then
                    'Hay diferencia entre el parque y el remito
                    txt = "La cantidad de EXTINTORES de préstamo que figuran en el parque del cliente no coincide con el remito."
                    MsgBox(txt)
                End If
            End If
            ddr.Close()

            'Consulto el parque de extintores de mangueras
            cm.Parameters.Add("macpdtcod_0", OracleType.VarChar).Value = ARTICULO_PRESTAMO_MAN
            ddr = cm.ExecuteReader

            If ddr.Read Then

                If CInt(ddr("macqty_0")) <> l_PrestamosMan Then
                    'Hay diferencia entre el parque y el remito
                    If txt <> "" Then txt &= vbCrLf & vbCrLf
                    txt &= "La cantidad de MANGUERAS de préstamo que figuran en el parque del cliente no coincide con el remito." & vbCrLf & vbCrLf
                    MsgBox(txt)
                End If
            End If
            ddr.Close()
            ddr.Dispose()

            If txt = "" Then
                Return True

            Else
                RaiseEvent ErrorDocumento(Me, New ErrDocumentoEvenArgs(txt))
                Return False

            End If

        Catch ex As Exception
            RaiseEvent ErrorDocumento(Me, New ErrDocumentoEvenArgs(ex.Message))
            Return False

        End Try

    End Function
    Private Function AnalizarRemito() As Boolean
        Dim dr As DataRow

        'Rutina de análisis de Remitos de nuevos
        For Each dr In dtRemitod.Rows

            If dr("tclcod_0").ToString = "10" Then
                l_Equipos += CType(dr("qty_0"), Integer)
                l_Peso += CType(dr("peso"), Double)
                l_Peso2 += CType(dr("peso"), Double)

            ElseIf dr("cfglin_0").ToString = "651" Then
                l_TieneInstalacion = True
                l_Instalaciones += CType(dr("qty_0"), Integer)

            Else
                l_Peso += CType(dr("peso"), Double)
                l_Peso2 += CType(dr("peso"), Double)

                l_Varios = True

            End If

        Next

        If dtRemitod.Rows.Count = 0 Then
            'Disparar evento "Este documento no contiene ítems de logistica"
            RaiseEvent ErrorDocumento(Me, New ErrDocumentoEvenArgs("Este documento no contiene ítems de logística"))
            Return False

        Else
            Return True

        End If

    End Function
    Public Function Buscar(ByVal Documento As String) As Boolean

        Documento = Documento.ToUpper.Trim
        l_NroDocumento = Documento

        l_Equipos = 0
        l_Mangas = 0
        l_PrestamosExt = 0
        l_PrestamosMan = 0
        l_RechazosMan = 0
        l_RechazosExt = 0
        l_Instalaciones = 0
        l_Varios = False
        l_Peso = 0
        l_Peso2 = 0
        l_TieneInstalacion = False
        l_EsTarea = False

        dtSolicitud.Rows.Clear()
        dtIntervencion.Rows.Clear()
        dtRetiros.Rows.Clear()
        dtConsumos.Rows.Clear()
        dtRemito.Rows.Clear()
        dtRemitod.Rows.Clear()

        If EsIntervencion Then

            daIntervencion.SelectCommand.Parameters(0).Value = Documento
            daIntervencion.Fill(dtIntervencion)

            If dtIntervencion.Rows.Count = 1 Then

                Select Case CInt(dtIntervencion.Rows(0).Item("zflgtrip_0"))
                    Case 7  'Intervencion A resolver
                        RaiseEvent ErrorDocumento(Me, New ErrDocumentoEvenArgs("No se puede poner en ruta una intervención en estado A Resolver"))

                    Case 8 'Intervención Cerrada
                        RaiseEvent ErrorDocumento(Me, New ErrDocumentoEvenArgs("No se puede poner en ruta una intervención en estado Cerrado"))

                    Case Else

                        daSolicitud.SelectCommand.Parameters(0).Value = dtIntervencion.Rows(0).Item("srvdemnum_0").ToString
                        daSolicitud.Fill(dtSolicitud)

                        daRetiros.SelectCommand.Parameters(0).Value = Documento
                        daRetiros.Fill(dtRetiros)

                        daConsumos.SelectCommand.Parameters(0).Value = Documento
                        daConsumos.Fill(dtConsumos)

                        If TieneRemito Then
                            Buscar = AnalizarIntervencionEntrega()

                        Else
                            Buscar = AnalizarIntervencionRetiro()

                        End If

                End Select

            Else
                RaiseEvent ErrorDocumento(Me, New ErrDocumentoEvenArgs("No se encontró la intervención"))
                Buscar = False

            End If

        Else
            daRemito.SelectCommand.Parameters(0).Value = Documento
            daRemitod.SelectCommand.Parameters(0).Value = Documento

            daRemito.Fill(dtRemito)
            daRemitod.Fill(dtRemitod)

            If dtRemito.Rows.Count = 1 Then
                Buscar = AnalizarRemito()

            Else
                RaiseEvent ErrorDocumento(Me, New ErrDocumentoEvenArgs("No se encontró el remito"))
                Buscar = False

            End If

        End If

    End Function
    Public Function BuscarAlRecibir(ByVal Documento As String) As Boolean
        Documento = Documento.ToUpper.Trim
        l_NroDocumento = Documento

        l_Equipos = 0
        l_Mangas = 0
        l_PrestamosExt = 0
        l_PrestamosMan = 0
        l_RechazosMan = 0
        l_RechazosExt = 0
        l_Instalaciones = 0
        l_Varios = False
        l_Peso = 0
        l_Peso2 = 0
        l_TieneInstalacion = False
        l_EsTarea = False

        dtSolicitud.Rows.Clear()
        dtIntervencion.Rows.Clear()
        dtRetiros.Rows.Clear()
        dtConsumos.Rows.Clear()
        dtRemito.Rows.Clear()
        dtRemitod.Rows.Clear()

        If EsIntervencion Then

            daIntervencion.SelectCommand.Parameters(0).Value = Documento
            daIntervencion.Fill(dtIntervencion)

            If dtIntervencion.Rows.Count = 1 Then


                daSolicitud.SelectCommand.Parameters(0).Value = dtIntervencion.Rows(0).Item("srvdemnum_0").ToString
                daSolicitud.Fill(dtSolicitud)

                daRetiros.SelectCommand.Parameters(0).Value = Documento
                daRetiros.Fill(dtRetiros)

                daConsumos.SelectCommand.Parameters(0).Value = Documento
                daConsumos.Fill(dtConsumos)

                If TieneRemito Then
                    BuscarAlRecibir = AnalizarIntervencionEntrega()
                End If
            End If
        End If
    End Function
    Public Function ValidarDiaEntrega(ByVal Fecha As Date) As Boolean
        Dim dr As OracleDataReader
        Dim cm As New OracleCommand("SELECT uvyday1_0, uvyday2_0, uvyday3_0, uvyday4_0, uvyday5_0, uvyday6_0, uvyday7_0 FROM bpdlvcust WHERE bpcnum_0 = :p1 AND bpaadd_0 = :p2", cn)
        cm.Parameters.Add("p1", OracleType.VarChar).Value = Tercero
        cm.Parameters.Add("p2", OracleType.VarChar).Value = Sucursal

        dr = cm.ExecuteReader(CommandBehavior.SingleRow)

        If dr.Read Then

            Select Case Fecha.DayOfWeek
                Case DayOfWeek.Monday
                    ValidarDiaEntrega = CType(IIf(CType(dr("uvyday1_0"), Integer) = 2, True, False), Boolean)

                Case DayOfWeek.Tuesday
                    ValidarDiaEntrega = CType(IIf(CType(dr("uvyday2_0"), Integer) = 2, True, False), Boolean)

                Case DayOfWeek.Wednesday
                    ValidarDiaEntrega = CType(IIf(CType(dr("uvyday3_0"), Integer) = 2, True, False), Boolean)

                Case DayOfWeek.Thursday
                    ValidarDiaEntrega = CType(IIf(CType(dr("uvyday4_0"), Integer) = 2, True, False), Boolean)

                Case DayOfWeek.Friday
                    ValidarDiaEntrega = CType(IIf(CType(dr("uvyday5_0"), Integer) = 2, True, False), Boolean)

                Case DayOfWeek.Saturday
                    ValidarDiaEntrega = CType(IIf(CType(dr("uvyday6_0"), Integer) = 2, True, False), Boolean)

                Case DayOfWeek.Sunday
                    ValidarDiaEntrega = CType(IIf(CType(dr("uvyday7_0"), Integer) = 2, True, False), Boolean)

            End Select

        Else
            ValidarDiaEntrega = False

        End If

        dr.Close() : dr = Nothing
        cm.Dispose() : cm = Nothing

    End Function

    Public ReadOnly Property Tipo() As String
        Get

            If EsIntervencion Then
                If TieneRemito OrElse TipoIntervencion = "G1" Then
                    Return "ENT"

                Else
                    Return IIf(l_EsTarea, "CTL", "RET").ToString

                End If

            Else
                If l_TieneInstalacion Then
                    If l_Equipos > 0 Or l_Varios Then
                        Return "NCI"

                    Else
                        Return "INS"

                    End If
                Else
                    Return "NUE"

                End If

            End If

        End Get
    End Property
    Public ReadOnly Property Documento() As String
        Get
            'If EsIntervencion Then
            '    Return dtIntervencion.Rows(0).Item("num_0").ToString
            'Else
            '    Return dtRemito.Rows(0).Item("sdhnum_0").ToString
            'End If
            Return l_NroDocumento
        End Get
    End Property
    Public ReadOnly Property RemitoIntervencion() As String
        Get
            If EsIntervencion Then
                Return dtIntervencion.Rows(0).Item("ysdhdeb_0").ToString

            Else
                Return " "

            End If
        End Get
    End Property
    Public ReadOnly Property PedidoRemito() As String
        Get
            If EsIntervencion Then
                Return " "

            Else
                Return dtRemito.Rows(0).Item("sohnum_0").ToString

            End If
        End Get
    End Property
    Public ReadOnly Property Tercero() As String
        Get
            If EsIntervencion Then
                Return dtIntervencion.Rows(0).Item("bpc_0").ToString
            Else
                Return dtRemito.Rows(0).Item("bpcord_0").ToString
            End If
        End Get
    End Property
    Public ReadOnly Property Nombre() As String
        Get
            If EsIntervencion Then
                Return NombreCliente(dtIntervencion.Rows(0).Item("bpc_0").ToString)
            Else
                Return NombreCliente(dtRemito.Rows(0).Item("bpcord_0").ToString)
            End If
        End Get
    End Property
    Public ReadOnly Property Sucursal() As String
        Get
            If EsIntervencion Then
                Return dtIntervencion.Rows(0).Item("bpaadd_0").ToString
            Else
                Return dtRemito.Rows(0).Item("bpaadd_0").ToString
            End If
        End Get
    End Property
    Public ReadOnly Property Domicilio() As String
        Get
            If EsIntervencion Then
                Return dtIntervencion.Rows(0).Item("add_0").ToString
            Else
                Return dtRemito.Rows(0).Item("bpdaddlig_0").ToString
            End If
        End Get
    End Property
    Public ReadOnly Property Localidad() As String
        Get
            If EsIntervencion Then
                Return dtIntervencion.Rows(0).Item("cty_0").ToString
            Else
                Return dtRemito.Rows(0).Item("bpdcty_0").ToString
            End If
        End Get
    End Property
    Public ReadOnly Property Equipos() As Integer
        Get
            Return l_Equipos
        End Get
    End Property
    Public ReadOnly Property Mangueras() As Integer
        Get
            Return l_Mangas
        End Get
    End Property
    Public ReadOnly Property PrestamosExtintores() As Integer
        Get
            Return l_PrestamosExt
        End Get
    End Property
    Public ReadOnly Property PrestamosMangueras() As Integer
        Get
            Return l_PrestamosMan
        End Get
    End Property
    Public ReadOnly Property Instalaciones() As Integer
        Get
            Return l_Instalaciones
        End Get
    End Property
    Public ReadOnly Property RechazosExtintor() As Integer
        Get
            Return l_RechazosExt
        End Get
    End Property
    Public ReadOnly Property RechazosManguera() As Integer
        Get
            Return l_RechazosMan
        End Get
    End Property
    Public ReadOnly Property Cobranza() As Boolean
        Get
            If EsIntervencion Then
                If TieneRemito Then
                    Select Case dtSolicitud.Rows(0).Item("srepte_0").ToString
                        Case "001", "002"
                            Return CInt(dtIntervencion.Rows(0).Item("dlvpio_0")) = 1

                        Case Else
                            Return False

                    End Select

                Else
                    Return False

                End If

            Else
                Select Case dtRemito.Rows(0).Item("pte_0").ToString
                    Case "001", "002"
                        Return True

                    Case Else
                        Return False

                End Select

            End If

        End Get
    End Property
    Public ReadOnly Property Varios() As Boolean
        Get
            Return l_Varios
        End Get
    End Property
    Public ReadOnly Property Peso() As Double
        Get
            Return l_Peso
        End Get
    End Property
    Public ReadOnly Property PesoUnigis() As Double
        Get
            Return l_Peso2
        End Get
    End Property
    Public ReadOnly Property Hora() As String
        Get
            If EsIntervencion Then
                Dim dr As DataRow = dtIntervencion.Rows(0)
                Dim l_Hora As String = " "

                If dr("yhdesde1_0").ToString <> "0000" And dr("yhhasta1_0").ToString <> "0000" Then
                    l_Hora = String.Format("{0} a {1}", dr("yhdesde1_0").ToString, dr("yhhasta1_0").ToString)

                    If dr("yhdesde2_0").ToString <> "0000" And dr("yhhasta2_0").ToString <> "0000" Then
                        l_Hora &= " y "
                        l_Hora &= String.Format("{0} a {1}", dr("yhdesde2_0").ToString, dr("yhhasta2_0").ToString)
                    End If

                End If

                Return l_Hora

            Else
                Return " "

            End If
        End Get
    End Property
    Public ReadOnly Property TieneRemito() As Boolean
        Get
            If EsIntervencion Then
                Return (dtIntervencion.Rows(0).Item("ysdhdeb_0").ToString <> " ")
            Else
                Return False
            End If
        End Get
    End Property
    Public ReadOnly Property TieneInstalacion() As Boolean
        Get
            Return l_TieneInstalacion
        End Get
    End Property
    Public ReadOnly Property EsTarea() As Boolean
        Get
            Return l_EsTarea
        End Get
    End Property
    Public ReadOnly Property ModoEntrega() As String
        Get
            If EsIntervencion Then
                Return dtIntervencion.Rows(0).Item("mdl_0").ToString
            Else
                Return dtRemito.Rows(0).Item("mdl_0").ToString
            End If
        End Get
    End Property
    Public ReadOnly Property FechaEntrega() As Date
        Get
            If EsIntervencion Then
                Return CType(dtIntervencion.Rows(0).Item("datend_0"), Date)
            Else
                Return CType(dtRemito.Rows(0).Item("dlvdat_0"), Date)
            End If
        End Get
    End Property
    Public ReadOnly Property Serie() As String
        Get
            Return l_Serie
        End Get
    End Property
    Public ReadOnly Property TipoIntervencion() As String
        Get
            Dim dr As DataRow

            If dtIntervencion.Rows.Count = 1 Then
                dr = dtIntervencion.Rows(0)
                Return dr("typ_0").ToString
            Else
                Return ""
            End If
        End Get
    End Property
    Public ReadOnly Property EsIntervencion() As Boolean
        Get
            EsIntervencion = (l_NroDocumento.IndexOf("RMR") = -1)
        End Get
    End Property
    Public ReadOnly Property Intervencion() As Intervencion
        Get
            Dim dr As DataRow = dtIntervencion.Rows(0)
            Dim itn As New Intervencion(cn)

            If itn.Abrir(dr("num_0").ToString) Then
                Return itn
            Else
                Return Nothing
            End If
        End Get
    End Property

End Class

Public Class ErrDocumentoEvenArgs
    Inherits System.EventArgs

    Public Mensaje As String = ""

    Public Sub New(ByVal Msg As String)
        Me.Mensaje = Msg
    End Sub

End Class