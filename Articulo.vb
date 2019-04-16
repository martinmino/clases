Imports System.Data.OracleClient

Public Class Articulo

    Implements IDisposable

    Private da As OracleDataAdapter 'itmmaster
    Private da2 As OracleDataAdapter 'itmsales
    Private dt As New DataTable
    Private dt2 As New DataTable
    Private dr As DataRow
    Private disposedValue As Boolean = False        ' Para detectar llamadas redundantes
    Private cn As OracleConnection

    'NEW
    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
        Adaptadores()
    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal Codigo As String)
        Me.New(cn)

        Abrir(Codigo)
    End Sub
    Private Sub Adaptadores()
        Dim Sql As String

        Sql = "SELECT * FROM itmmaster WHERE itmref_0 = :itmref"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("itmref", OracleType.VarChar)

        Sql = "SELECT * FROM itmsales WHERE itmref_0 = :itmref"
        da2 = New OracleDataAdapter(Sql, cn)
        da2.SelectCommand.Parameters.Add("itmref", OracleType.VarChar)

    End Sub
    Public Function Abrir(ByVal Codigo As String) As Boolean

        da.SelectCommand.Parameters("itmref").Value = Codigo
        da2.SelectCommand.Parameters("itmref").Value = Codigo

        dt.Clear()
        dt2.Clear()
        da.Fill(dt)
        da2.Fill(dt2)

        If dt.Rows.Count = 1 Then dr = dt.Rows(0)

        Return dt.Rows.Count <> 0

    End Function
    Shared Function Tabla(ByVal cn As OracleConnection) As DataTable
        Dim Sql As String
        Dim dt As New DataTable
        Dim da As OracleDataAdapter

        Sql = "SELECT itmref_0, itmdes1_0 FROM itmmaster ORDER BY itmref_0"
        da = New OracleDataAdapter(Sql, cn)
        da.Fill(dt)
        da.Dispose()

        Return dt

    End Function
    Shared Function TablaConPrecio(ByVal cn As OracleConnection) As DataTable
        Dim Sql As String
        Dim dt As New DataTable
        Dim da As OracleDataAdapter

        Sql = "SELECT xpe.itmref_0, itmdes1_0, (xpe.itmref_0 || ' - ' || itmdes1_0)  as completo FROM xprecios xpe inner join itmmaster itm on (xpe.itmref_0 = itm.itmref_0) order by itmref_0 asc"
        da = New OracleDataAdapter(Sql, cn)
        da.Fill(dt)
        da.Dispose()

        Return dt

    End Function
    Public Function Costo() As Double
        Dim Sql As String
        Dim dt As New DataTable
        Dim da As OracleDataAdapter
        Dim dr As DataRow
        Dim i As Double = 0

        Sql = "SELECT itmref_0, vlttot_0 "
        Sql &= "FROM itmcost "
        Sql &= "WHERE csttyp_0 = 1 AND "
        Sql &= "      stofcy_0 = 'D01'  AND "
        Sql &= "      itmref_0 = :itmref "
        Sql &= "ORDER BY yea_0 DESC, itcdat_0 DESC"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("itmref", OracleType.VarChar).Value = Codigo
        da.Fill(dt)
        da.Dispose()

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            i = CDbl(dr("vlttot_0"))
        End If

        Return i

    End Function
    Public Function ArticuloParaPedido() As Boolean
        'TRUE si el articulo se permite en pedidos, sino FALSE

        Dim Sql As String
        Dim da As OracleDataAdapter
        Dim dt As New DataTable

        Sql = "SELECT * FROM xprecio2 WHERE itmref_0 = :itmref AND ped_0 = 2"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("itmref", OracleType.VarChar).Value = Codigo
        da.Fill(dt)
        da.Dispose()

        Return dt.Rows.Count > 0

    End Function

    'PROPERTY
    Public ReadOnly Property Categoria() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("tclcod_0").ToString
        End Get
    End Property
    Public ReadOnly Property Descripcion() As String
        Get
            If Not IsNothing(dr) Then
                Return dr("itmdes1_0").ToString
            Else
                Return ""
            End If
        End Get
    End Property
    Public ReadOnly Property Codigo() As String
        Get
            If Not IsNothing(dr) Then
                Return dr("itmref_0").ToString
            Else
                Return ""
            End If
        End Get
    End Property
    Public ReadOnly Property Impuesto(ByVal bpc As Cliente) As Impuesto
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return New Impuesto(cn, dr("vacitm_0").ToString, bpc)
        End Get
    End Property
    Public ReadOnly Property Unidad() As String
        Get
            If Not IsNothing(dr) Then
                Return dr("stu_0").ToString
            Else
                Return ""
            End If
        End Get
    End Property
    Public ReadOnly Property UnidadVta() As String
        Get
            If Not IsNothing(dr) Then
                Return dr("sau_0").ToString
            Else
                Return ""
            End If
        End Get
    End Property
    Public ReadOnly Property GestionSerie() As Integer
        Get
            If Not IsNothing(dr) Then
                Return CInt(dr("sermgtcod_0"))
            Else
                Return 0
            End If
        End Get
    End Property
    Public ReadOnly Property GestionStock() As Integer
        Get
            If Not IsNothing(dr) Then
                Return CInt(dr("stomgtcod_0"))
            Else
                Return 0
            End If
        End Get
    End Property
    Public ReadOnly Property ArtParaParque() As Boolean
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return CInt(dr("xparque_0")) = 2

            Else
                Return False

            End If

        End Get
    End Property
    Public ReadOnly Property Grupo() As String
        Get
            If Not IsNothing(dr) Then
                Return dr("xgrp_0").ToString.Trim
            Else
                Return ""
            End If
        End Get
    End Property
    Public ReadOnly Property LineaProducto() As String
        Get
            If Not IsNothing(dr) Then
                Return dr("cfglin_0").ToString.Trim
            Else
                Return ""
            End If
        End Get
    End Property
    Public ReadOnly Property Familia(ByVal Indice As Integer) As String
        Get
            If Not IsNothing(dr) Then
                Return dr("tsicod_" & Indice.ToString).ToString.Trim
            Else
                Return " "
            End If
        End Get
    End Property
    Public ReadOnly Property FamiliaDescripcion(ByVal Indice As Integer) As String
        Get
            Dim Descripcion As String = ""
            Dim NroTabla(4) As String
            Dim Sql As String
            Dim da As OracleDataAdapter
            Dim dt As New DataTable

            NroTabla(0) = "20"
            NroTabla(1) = "21"
            NroTabla(2) = "22"
            NroTabla(3) = "23"
            NroTabla(4) = "24"

            Sql = "select * "
            Sql &= "from atextra "
            Sql &= "where codfic_0 = 'ATABDIV' and "
            Sql &= "	  zone_0 = 'LNGDES' and "
            Sql &= "	  langue_0 = 'SPA' and "
            Sql &= "	  ident1_0 = :tabla and "
            Sql &= "	  ident2_0 = :codigo "

            da = New OracleDataAdapter(Sql, cn)
            da.SelectCommand.Parameters.Add("tabla", OracleType.VarChar).Value = NroTabla(Indice)
            da.SelectCommand.Parameters.Add("codigo", OracleType.VarChar).Value = Familia(Indice).Trim
            da.Fill(dt)
            da.Dispose()

            If dt.Rows.Count > 0 Then
                Dim dr As DataRow
                dr = dt.Rows(0)

                Descripcion = dr("texte_0").ToString

            End If

            Return Descripcion

        End Get
    End Property
    Public ReadOnly Property EjeAnalitico(ByVal Indice As Integer) As String
        Get
            If Not IsNothing(dr) Then
                Return dr("cce_" & Indice.ToString).ToString
            Else
                Return " "
            End If
        End Get
    End Property
    Public ReadOnly Property peso() As Double
        Get
            If Not IsNothing(dr) Then
                Return CDbl(dr("itmwei_0"))
            Else
                Return 0
            End If
        End Get
    End Property
    Public ReadOnly Property AutoAsignacion() As Boolean
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CBool(IIf(dr("xautoasig_0").ToString = "2", True, False))
        End Get
    End Property
    Public ReadOnly Property Activo() As Boolean
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CBool(IIf(dr("itmsta_0").ToString = "1", True, False))
        End Get
    End Property
    Public ReadOnly Property EsCarro() As Boolean
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CBool(IIf(dr("xcarro_0").ToString = "2", True, False))
        End Get
    End Property

    Public ReadOnly Property LlevaIram() As Boolean
        Get
            Dim dr As DataRow
            Dim flg As Boolean = False

            If dt2.Rows.Count > 0 Then
                dr = dt2.Rows(0)
                flg = CInt(dr("yflgiram_0")) = 2
            End If

            Return flg
        End Get
    End Property
    Public ReadOnly Property LlevaTarjeta() As Boolean
        Get
            Dim dr As DataRow
            Dim flg As Boolean = False

            If dt2.Rows.Count > 0 Then
                dr = dt2.Rows(0)
                flg = CInt(dr("yflgsat_0")) = 2
            End If

            Return flg
        End Get
    End Property
    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If disposedValue Then Exit Sub

        If disposing Then
            ' TODO: Liberar otro estado (objetos administrados).
        End If

        ' TODO: Liberar su propio estado (objetos no administrados).
        ' TODO: Establecer campos grandes como Null.
        'da.Dispose()
        'dt.Dispose()

        Me.disposedValue = True
    End Sub

    ' Visual Basic agregó este código para implementar correctamente el modelo descartable.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' No cambie este código. Coloque el código de limpieza en Dispose (ByVal que se dispone como Boolean).
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
        Dispose(False)
    End Sub

End Class