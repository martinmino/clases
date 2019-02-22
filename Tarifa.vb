Imports System.Data.OracleClient

Public Class Tarifa
    Private cn As OracleConnection
    Private Listas() As String
    Private dt As New DataTable 'Todas las listas de precios vigentes
    Private dv As DataView
    Private dap As OracleDataAdapter
    Private dtPrecios As New DataTable
    Private TipoCambio As Double = 1

    Private ListaDistribuidor(6) As Integer
    Private ListaEmpresa(6) As Integer
    Private ListaConsorcios(6) As Integer
    Private ListaDirecta(3) As Integer

    Public Sub New(ByVal cn As OracleConnection)
        Dim da As OracleDataAdapter
        Dim Sql As String
        Dim dt2 As New DataTable

        Me.cn = cn

        'Recupera todas las listas de precios ordenadas por prioridad
        Sql = "SELECT pli_0, pio_0 "
        Sql &= "FROM spricconf "
        Sql &= "WHERE plienaflg_0 = 2 "
        Sql &= "ORDER BY pio_0"

        da = New OracleDataAdapter(Sql, cn)
        da.Fill(dt2)
        da.Dispose()

        ReDim Listas(dt2.Rows.Count - 1)

        For i = 0 To dt2.Rows.Count - 1
            Dim dr As DataRow = dt2.Rows(i)
            Listas(i) = dr(0).ToString

            'If dr(0).ToString = "TAR001" Then
            '    i = i + 1
            'End If

        Next

        'Listas(8) = "TAR002"

        dt2.Dispose()
        dt2 = Nothing

        'Recupero toda las listas vigentes con los precios 
        Sql = "SELECT spl.* "
        Sql &= "FROM spriclist spl INNER JOIN spricconf spc ON (spl.pli_0 = spc.pli_0) "
        Sql &= "WHERE spl.plistrdat_0 <= :dat and "
        Sql &= "	  spl.plienddat_0 >= :dat and "
        Sql &= "	  spc.plienaflg_0 = 2 "
        Sql &= "ORDER BY spl.pli_0, spl.plicrd_0, spl.plilin_0"

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("dat", OracleType.DateTime).Value = Today
        da.Fill(dt)
        da.Dispose()

        dv = New DataView(dt)


        '05.09.2018 Obtención de ultimo tipo de cambio cargado en ADONIX
        Dim dtp As New DataTable

        Sql = "SELECT * "
        Sql &= "FROM tabchange "
        Sql &= "WHERE cur_0 = 'USD' AND "
        Sql &= "	  chgtyp_0 = 1 "
        Sql &= "ORDER BY upddat_0 DESC"
        da = New OracleDataAdapter(Sql, cn)
        da.Fill(dtp)

        If dtp.Rows.Count > 0 Then
            Dim dr As DataRow = dtp.Rows(0)
            TipoCambio = CDbl(dr("chgrat_0"))
        End If
        dtp.Dispose()
        da.Dispose()

        'Precios
        Sql = "SELECT * FROM xprecio2 WHERE itmref_0 = :itmref ORDER BY qty_0 DESC"
        dap = New OracleDataAdapter(Sql, cn)
        dap.SelectCommand.Parameters.Add("itmref", OracleType.VarChar)


        'Cada elemento de las listas le asigno el index el campo PRECIO_x <--- x=1
        ListaDistribuidor(0) = 0
        ListaDistribuidor(1) = 1
        ListaDistribuidor(2) = 2
        ListaDistribuidor(3) = 3
        ListaDistribuidor(4) = 4
        ListaDistribuidor(5) = 5
        ListaDistribuidor(6) = 5

        ListaEmpresa(0) = 10
        ListaEmpresa(1) = 6
        ListaEmpresa(2) = 7
        ListaEmpresa(3) = 8
        ListaEmpresa(4) = 9
        ListaEmpresa(5) = 10
        ListaEmpresa(6) = 10

        ListaConsorcios(0) = 15
        ListaConsorcios(1) = 11
        ListaConsorcios(2) = 12
        ListaConsorcios(3) = 13
        ListaConsorcios(4) = 14
        ListaConsorcios(5) = 15
        ListaConsorcios(6) = 15

        ListaDirecta(0) = 18
        ListaDirecta(1) = 16
        ListaDirecta(2) = 17
        ListaDirecta(3) = 18

    End Sub
    Public Function ObtenerPrecio(ByVal bpc As Cliente, ByVal itm As Articulo, Optional ByVal cantidad As Double = 0) As Double
        Dim Precio As Double = 0
        Dim i As Integer
        Dim Filtro As String = ""
        Dim HayPrecioAdonix As Boolean = False

        For i = 0 To Listas.Length - 1

            Filtro = "pli_0 = '" & Listas(i) & "' AND "

            Select Case Listas(i)
                Case "TAR001" 'Lista precios
                    Filtro &= "plicri1_0 = '" & itm.Codigo & "'"

                Case "TAR004" 'Lista de precios especiales
                    Filtro &= "plicri1_0 = '" & bpc.Codigo & "'"
                    Filtro &= " AND "
                    Filtro &= "plicri2_0 = '" & itm.Codigo & "'"

                Case "TAR005" 'Precios fijos tarj-estamp-etc
                    Filtro &= "plicri1_0 = '" & itm.Codigo & "'"
                    Filtro &= " AND "
                    Filtro &= "plicri2_0 = '" & bpc.Familia2 & "'"

                Case "TAR006" 'Abonados en $0
                    Filtro &= "plicri1_0 = '" & IIf(bpc.EsAbonado, 2, 1).ToString & "'"
                    Filtro &= " AND "
                    Filtro &= "plicri2_0 = '" & itm.Codigo & "'"

                Case "TAR007" 'Lista precios por vendedor
                    Filtro &= "plicri1_0 = '" & bpc.Representante(0) & "'"
                    Filtro &= " AND "
                    Filtro &= "plicri2_0 = '" & itm.Codigo & "'"

                Case Else
                    Continue For

            End Select

            dv.RowFilter = Filtro

            If dv.Count > 0 Then
                Dim dr As DataRow = dv.Item(0).Row
                Precio = CDbl(dr("pri_0"))
                HayPrecioAdonix = True
                Exit For
            End If
        Next

        If Precio = 0 AndAlso HayPrecioAdonix = False Then
            dtprecios.Clear()

            dap.SelectCommand.Parameters("itmref").Value = itm.Codigo
            dap.Fill(dtPrecios)

            If dtprecios.Rows.Count = 0 Then Return Precio

            Dim dr As DataRow = Nothing

            For x = 0 To dt.Rows.Count - 1
                If cantidad = 0 Then Return Precio

                If cantidad >= CDbl(dtprecios.Rows(x).Item("qty_0")) Then
                    dr = dtprecios.Rows(x)
                    Exit For
                End If
            Next

            'Select Case bpc.Familia2
            '    Case "10", "11"
            '        Precio = CDbl(dr("precio_0"))

            '    Case "20", "30", "70", "80", "90"
            '        Precio = CDbl(dr("precio_1"))

            '    Case "40", "50", "60"
            '        Precio = CDbl(dr("precio_2"))
            'End Select

            Dim col As String = SelectorDeColumna(bpc)

            Precio = CDbl(dr("precio_" & col))

            '05.09.2018 Si el precio es en dolares, se multiplica por el tipo de cambio
            'Se excluyen articulos 705020 y 705032
            'If CInt(dr("ped_0")) = 2 AndAlso itm.Codigo <> "705020" AndAlso itm.Codigo <> "705032" Then
            '    Precio = Precio * TipoCambio
            'End If

        End If

        Return Precio

    End Function
    Public Function ObtenerPrecio(ByVal Cliente As String, ByVal Articulo As String, Optional ByVal cantidad As Double = 0) As Double
        Dim itm As New Articulo(cn)
        Dim bpc As New Cliente(cn)

        bpc.Abrir(Cliente)

        If itm.Abrir(Articulo) Then
            Return ObtenerPrecio(bpc, itm, cantidad)

        Else
            Return 0

        End If

    End Function
    Public Function ObtenerPrecio(ByVal bpc As Cliente, ByVal Articulo As String, Optional ByVal cantidad As Double = 0) As Double
        Dim itm As New Articulo(cn)

        If itm.Abrir(Articulo) Then
            Return ObtenerPrecio(bpc, itm, cantidad)
        Else
            Return 0
        End If
    End Function
    Public Function ObtenerPrecio(ByVal bpc As Cliente, ByVal itm As Articulo, ByVal Planta As String) As Double
        Return ObtenerPrecio(bpc, itm.Codigo, Planta)
    End Function
    Public Function ObtenerPrecio(ByVal bpc As Cliente, ByVal itm As String, ByVal Planta As String) As Double
        Dim Precio As Double = 0
        Dim i As Integer
        Dim Filtro As String = ""

        For i = 0 To Listas.Length - 1

            Filtro = "pli_0 = '" & Listas(i) & "' AND "

            Select Case Listas(i)
                Case "TAR001" 'Lista precios
                    Filtro &= "plicri1_0 = '" & itm & "'"

                Case "TAR004" 'Lista de precios especiales
                    Filtro &= "plicri1_0 = '" & bpc.Codigo & "'"
                    Filtro &= " AND "
                    Filtro &= "plicri2_0 = '" & itm & "'"

                Case "TAR005" 'Precios fijos tarj-estamp-etc
                    Filtro &= "plicri1_0 = '" & itm & "'"
                    Filtro &= " AND "
                    Filtro &= "plicri2_0 = '" & bpc.Familia2 & "'"

                Case "TAR006" 'Abonados en $0
                    Filtro &= "plicri1_0 = '" & IIf(bpc.EsAbonado, 2, 1).ToString & "'"
                    Filtro &= " AND "
                    Filtro &= "plicri2_0 = '" & itm & "'"

                Case "TAR007"
                    Filtro &= "plicri1_0 = '" & Planta & "'"
                    Filtro &= " AND "
                    Filtro &= "plicri2_0 = '" & itm & "'"

                Case Else
                    Continue For

            End Select

            dv.RowFilter = Filtro

            If dv.Count > 0 Then
                Dim dr As DataRow = dv.Item(0).Row

                Precio = CDbl(dr("pri_0"))
                Exit For

            End If

        Next

        Return Precio

    End Function

    Public Function SelectorDeColumna(ByVal bpc As Cliente) As String
        Dim i As Integer = 0
        Dim n As Integer = 17 'Numero de columna
        Dim bpr As Cliente 'Tercero pagador

        If bpc.EsProspecto Then
            bpr = bpc
        Else
            bpr = bpc.TerceroPagador
        End If


        'Obtengo la categoría ABC
        i = bpr.TipoAbc

        'Si ABC es A, miro si tiene tilde de Plus
        If i = 0 Then i = 5
        If i = 1 AndAlso bpr.PlusAbc Then i = 0


        Select Case bpr.Familia2
            Case "20", "30", "80" 'Distribuidor/Constructora
                n = ListaDistribuidor(i)

            Case "40", "60", "70" 'Empresa
                n = ListaEmpresa(i)

            Case "10", "11", "50" 'Consorcios / Institucional / Licitaciones
                n = ListaConsorcios(i)

            Case "90" 'Empleados
                n = ListaDistribuidor(0)

            Case Else 'Directa

                n = ListaEmpresa(i)

        End Select

        Return n.ToString

    End Function
    Public Function EsNuevo(ByVal Articulo As String) As Boolean
        Dim dt As New DataTable
        Dim f As Boolean = False

        dap.SelectCommand.Parameters("itmref").Value = Articulo
        dap.Fill(dt)

        For Each dr As DataRow In dt.Rows
            If CInt(dr("ped_0")) = 2 Then f = True
            If f Then Exit For
        Next

        Return f

    End Function
    Public Function ExisteArticulo(ByVal Articulo As String) As Boolean
        Dim dt As New DataTable

        dap.SelectCommand.Parameters("itmref").Value = Articulo
        dap.Fill(dt)

        Return dt.Rows.Count > 0
    End Function

    Public ReadOnly Property CotizacionDolar() As Double
        Get
            Return TipoCambio
        End Get
    End Property

End Class