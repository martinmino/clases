Imports System.Data.OracleClient

Public Class Tercero
    Implements IDisposable

    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private dt As New DataTable
    Private disposedValue As Boolean = False        ' Para detectar llamadas redundantes

    'NEW
    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
        Adaptadores()
    End Sub

    'SUB
    Private Sub Adaptadores()
        Dim Sql As String

        Sql = "SELECT * FROM bpartner WHERE bprnum_0 = :bprnum_0"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("bprnum_0", OracleType.VarChar)
        da.InsertCommand = New OracleCommandBuilder(da).GetInsertCommand
        da.UpdateCommand = New OracleCommandBuilder(da).GetUpdateCommand

        da.FillSchema(dt, SchemaType.Mapped)

    End Sub
    Public Function Abrir(ByVal Codigo As String) As Boolean
        da.SelectCommand.Parameters("bprnum_0").Value = Codigo
        dt.Clear()
        da.Fill(dt)

        Return dt.Rows.Count > 0
    End Function
    Public Overloads Sub Grabar()
        da.Update(dt)
    End Sub
    Public Function ExisteSucursal(ByVal NroSuc As String) As Boolean
        Dim da = New OracleDataAdapter("SELECT * FROM bpaddress WHERE bpanum_0 = :bpanum_0 AND bpaadd_0 = :bpaadd_0", cn)
        Dim dt As New DataTable
        Dim f As Boolean = False

        da.SelectCommand.Parameters.Add("bpanum_0", OracleType.VarChar).Value = Codigo
        da.SelectCommand.Parameters.Add("bpaadd_0", OracleType.VarChar).Value = NroSuc

        Try
            da.Fill(dt)
            da.Dispose()

            f = dt.Rows.Count = 1

        Catch ex As Exception

        End Try

        Return f

    End Function

    Public Sub Nuevo() '(ByVal numero As String, ByVal terceros As String, ByVal nombre As String, ByVal fantasia As String, ByVal tip As String, ByVal doctip As String, ByVal dni As String, ByVal rep As String, ByVal fs1 As String, ByVal fs2 As String, ByVal fs3 As String, ByVal usuario As String, ByVal cant As Integer, ByVal pte As String, ByVal comision As String, ByVal iva As String, ByVal mailfc As String, ByVal observa As String)
        Dim dr As DataRow

        dr = dt.NewRow

        dr("BPRNUM_0") = " " 'numero
        dr("BPRNAM_0") = " " 'nombre
        dr("BPRNAM_1") = " "
        dr("BPRSHO_0") = " "
        dr("EECNUM_0") = " "
        dr("BETFCY_0") = 1
        dr("FCY_0") = " "
        dr("CRY_0") = "AR"
        dr("DOCTYP_0") = " " 'doctip
        dr("DOCNUM_0") = " " 'dni
        dr("CRN_0") = " " 'dni
        dr("NAF_0") = " "
        dr("CUR_0") = "ARS"
        dr("LAN_0") = "SPA"
        dr("BPRLOG_0") = " "
        dr("VATNUM_0") = " "
        dr("FISCOD_0") = " "
        dr("GRUGPY_0") = " "
        dr("GRUCOD_0") = " "
        dr("BPCFLG_0") = 2
        dr("BPSFLG_0") = 1
        dr("BPTFLG_0") = 1
        dr("FCTFLG_0") = 1
        dr("REPFLG_0") = 1
        dr("BPRACC_0") = 1
        dr("PPTFLG_0") = 1
        dr("PRVFLG_0") = 1
        dr("DOOFLG_0") = 1
        dr("PTHFLG_0") = 0
        dr("ACCCOD_0") = " "
        dr("BPRFLG_0") = 1
        dr("BPRFLG_1") = 1
        dr("BPRFLG_2") = 1
        dr("BPRFLG_3") = 1
        dr("BPAADD_0") = " " '"001"
        dr("CNTNAM_0") = " "
        dr("BIDNUM_0") = " "
        dr("BIDCRY_0") = " "
        dr("SCT1_0") = " "
        dr("SCT2_0") = " "
        dr("ACS_0") = " "
        dr("EXPNUM_0") = 0
        dr("BPRGTETYP_0") = " "
        dr("CREUSR_0") = " "
        dr("CREDAT_0") = Date.Today 'Today.Hour * 3600 + Today.Minute * 60
        dr("UPDUSR_0") = " "
        dr("UPDDAT_0") = #12/31/1599# 'Date.Today 'Today.Hour * 3600 + Today.Minute * 60
        dr("XIIBBNUM_0") = " "

        dt.Rows.Add(dr)

    End Sub
    Public Sub SetSucursalPrincipal(ByVal Suc As String)
        Dim dr As DataRow

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            dr.BeginEdit()
            dr("bpaadd_0") = Suc
            dr.EndEdit()
        End If
    End Sub
    'PROPERTY
    Public Property Codigo() As String
        Get
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0).Item("bprnum_0").ToString
            Else
                Return ""
            End If

        End Get
        Set(ByVal value As String)
            If dt.Rows.Count > 0 Then
                Dim dr As DataRow

                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("bprnum_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property

    Public ReadOnly Property SucursalDefault() As Sucursal
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return New Sucursal(cn, dr("bprnum_0").ToString, dr("bpaadd_0").ToString)
        End Get
    End Property

    Public ReadOnly Property Divisa() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("cur_0").ToString
        End Get
    End Property
    Public ReadOnly Property Idioma() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("lan_0").ToString
        End Get
    End Property
    Public ReadOnly Property Contacto() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("cntnam_0").ToString
        End Get
    End Property
    Public ReadOnly Property Sucursal(ByVal NroSucursal As String) As Sucursal
        Get
            Return New Sucursal(cn, Codigo, NroSucursal)
        End Get
    End Property
    Public Property Nombre() As String
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return dr("bprnam_0").ToString
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow = dt.Rows(0)
            dr.BeginEdit()
            dr("bprnam_0") = IIf(value.Trim = "", " ", value.ToUpper)
            dr.EndEdit()
        End Set
    End Property
    Public ReadOnly Property EsProveedor() As Boolean
        Get
            Dim da As OracleDataAdapter
            Dim Sql As String = "SELECT * FROM bpsupplier WHERE bpsnum_0 = :bpsnum_0"
            Dim dt As New DataTable

            da = New OracleDataAdapter(Sql, cn)
            da.SelectCommand.Parameters.Add("bpsnum_0", OracleType.VarChar).Value = Codigo

            da.Fill(dt)
            da.Dispose()

            Return (dt.Rows.Count = 1)

        End Get
    End Property
    Public Property CUIT() As String
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return dr("docnum_0").ToString.Trim
            Else
                Return ""
            End If
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("docnum_0") = IIf(value.Trim = "", " ", value.ToUpper)
                dr("crn_0") = IIf(value.Trim = "", " ", value.ToUpper)
                dr.EndEdit()
            End If
        End Set
    End Property
    Public Property TipoDoc() As String
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return dr("doctyp_0").ToString
            Else
                Return " "
            End If
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("doctyp_0") = IIf(value.Trim = "", " ", value.ToUpper)
                dr.EndEdit()
            End If
        End Set
    End Property
    Public ReadOnly Property Cliente() As Cliente
        Get
            Dim bpc As New Cliente(cn)

            If bpc.Abrir(Codigo) Then
                Return bpc
            Else
                Return Nothing
            End If

        End Get
    End Property
    Public ReadOnly Property Proveedor() As Proveedor
        Get
            Dim bps As New Proveedor(cn)

            If bps.Abrir(Codigo) Then
                Return bps
            Else
                Return Nothing
            End If
        End Get
    End Property
    Public Property EsCliente() As Boolean
        Get
            Dim dr As DataRow
            Dim flg As Boolean = False

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)

                flg = CInt(dr("bpcflg_0")) = 2
            End If

            Return flg

        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("bpcflg_0") = IIf(value, 2, 1)
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property EsProspecto() As Boolean
        Get
            Dim dr As DataRow
            Dim flg As Boolean = False

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)

                flg = CInt(dr("pptflg_0")) = 2
            End If

            Return flg

        End Get
        Set(ByVal value As Boolean)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("pptflg_0") = IIf(value, 2, 1)
                dr.EndEdit()
            End If

        End Set
    End Property
    Public Property IIBB() As String
        Get
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Return dr("xiibbnum_0").ToString.Trim
            Else
                Return ""
            End If
        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("xiibbnum_0") = IIf(value.Trim = "", " ", value.ToUpper)
                dr.EndEdit()
            End If
        End Set
    End Property

    ' IDisposable
    ' Visual Basic agregó este código para implementar correctamente el modelo descartable.
    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        ' No cambie este código. Coloque el código de limpieza en Dispose (ByVal que se dispone como Boolean).
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
    Protected Overridable Overloads Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: Liberar otro estado (objetos administrados).
            End If

            ' TODO: Liberar su propio estado (objetos no administrados).
            ' TODO: Establecer campos grandes como Null.
            da.Dispose()
            dt.Dispose()
        End If

        Me.disposedValue = True

    End Sub
    Protected Overrides Sub Finalize()
        Dispose(False)
    End Sub

End Class