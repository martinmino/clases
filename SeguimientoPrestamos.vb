Imports System.Data.OracleClient

Public Class SeguimientoPrestamos
    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable

    'NEW
    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
        Adaptadores()
    End Sub

    Private Sub Adaptadores()
        Dim Sql As String

        Sql = "SELECT * FROM xprestamos WHERE itn_0 = :itn_0"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("itn_0", OracleType.VarChar)

        da.InsertCommand = New OracleCommandBuilder(da).GetInsertCommand
        da.UpdateCommand = New OracleCommandBuilder(da).GetUpdateCommand
        da.DeleteCommand = New OracleCommandBuilder(da).GetDeleteCommand
    End Sub

    Public Function Abrir(ByVal itn As String) As Boolean
        da.SelectCommand.Parameters("itn_0").Value = itn

        If dt Is Nothing Then
            dt = New DataTable
        Else
            dt.Clear()
        End If

        da.Fill(dt)

        Return dt.Rows.Count > 0

    End Function
    Public Function Nuevo() As Boolean
        Dim dr As DataRow

        If dt Is Nothing Then
            dt = New DataTable
            da.FillSchema(dt, SchemaType.Source)

        Else
            dt.Clear()
            dr = dt.NewRow
            dr("fecha_0") = Date.Today
            dr("itn_0") = " "
            dr("ruta_0") = 0
            dr("faltante_0") = 0
            dr("faltante_1") = 0
            dr("retencion_0") = 0
            dr("retencion_1") = 0
            dr("recuperado_0") = 0
            dr("recuperado_1") = 0
            dr("facturado_0") = 0
            dr("facturado_1") = 0
            dr("ok_0") = 0
            dr("ok_1") = 0
            dr("convertido_0") = 0
            dr("convertido_1") = 0

            dt.Rows.Add(dr)
        End If

    End Function
    Public Function Grabar() As Boolean
        Try
            da.Update(dt)
            Return True

        Catch ex As Exception
            Return False

        End Try
    End Function

    Public Property Fecha() As Date
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return CDate(dr("fecha_0"))
            Else
                Return #12/31/1599#
            End If

        End Get
        Set(ByVal value As Date)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("fecha_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property NumeroIntervencion() As String
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return dr("vcrnum_0").ToString
            Else
                Return " "
            End If

        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("itn_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property NumeroRuta() As Long
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return CLng(dr("ruta_0"))
            Else
                Return 0
            End If

        End Get
        Set(ByVal value As Long)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("ruta_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property FaltanteExtintor() As Long
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return CLng(dr("faltante_0"))
            Else
                Return 0
            End If

        End Get
        Set(ByVal value As Long)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("faltante_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property FaltanteManguera() As Long
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return CLng(dr("faltante_1"))
            Else
                Return 0
            End If

        End Get
        Set(ByVal value As Long)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("faltante_1") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property RetencionExtintor() As Long
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return CLng(dr("retencion_0"))
            Else
                Return 0
            End If

        End Get
        Set(ByVal value As Long)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("retencion_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property RetencionManguera() As Long
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return CLng(dr("retencion_1"))
            Else
                Return 0
            End If

        End Get
        Set(ByVal value As Long)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("retencion_1") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property RecuperoExtintor() As Long
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return CLng(dr("recuperado_0"))
            Else
                Return 0
            End If

        End Get
        Set(ByVal value As Long)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("recuperado_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property RecuperoManguera() As Long
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return CLng(dr("recuperado_1"))
            Else
                Return 0
            End If

        End Get
        Set(ByVal value As Long)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("recuperado_1") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property FacturadoExtintor() As Long
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return CLng(dr("facturado_0"))
            Else
                Return 0
            End If

        End Get
        Set(ByVal value As Long)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("facturado_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property FacturadoManguera() As Long
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return CLng(dr("facturado_1"))
            Else
                Return 0
            End If

        End Get
        Set(ByVal value As Long)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("facturado_1") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property OkExtintor() As Long
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return CLng(dr("ok_0"))
            Else
                Return 0
            End If

        End Get
        Set(ByVal value As Long)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("ok_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property OkManguera() As Long
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return CLng(dr("ok_1"))
            Else
                Return 0
            End If

        End Get
        Set(ByVal value As Long)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("ok_1") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property ConvertidoExtintor() As Long
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return CLng(dr("convertido_0"))
            Else
                Return 0
            End If

        End Get
        Set(ByVal value As Long)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("convertido_0") = value
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property ConvertidoManguera() As Long
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return CLng(dr("convertido_1"))
            Else
                Return 0
            End If

        End Get
        Set(ByVal value As Long)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("convertido_1") = value
                dr.EndEdit()
            End If

        End Set
    End Property

    Public ReadOnly Property ExtintoresFaltantes() As Long
        Get
            Dim dr As DataRow = dt.Rows(0)
            Dim i As Long = CLng(dr("faltante_0"))

            i -= CLng(dr("recuperado_0"))
            i -= CLng(dr("facturado_0"))
            i -= CLng(dr("ok_0"))
            i -= CLng(dr("convertido_0"))

            Return i
        End Get
    End Property
    Public ReadOnly Property ManguerasFaltantes() As Long
        Get
            Dim dr As DataRow = dt.Rows(0)
            Dim i As Long = CLng(dr("faltante_1"))

            i -= CLng(dr("recuperado_1"))
            i -= CLng(dr("facturado_1"))
            i -= CLng(dr("ok_1"))
            i -= CLng(dr("convertido_1"))

            Return i
        End Get
    End Property

End Class