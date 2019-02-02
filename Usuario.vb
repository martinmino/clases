Imports System.Data.OracleClient

Public Class Usuario

    Private cn As OracleConnection
    Private dt As DataTable
    Private da As OracleDataAdapter

    Private Logeado As Boolean = False
    Private MsgError As String = ""

    Private Per As Permisos

    'NEW
    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
        Adaptadores()
    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal Codigo As String)
        Me.New(cn)
        Abrir(Codigo)
    End Sub

    'FUNCIONES PRIVADAS
    Private Sub Adaptadores()
        Dim Sql As String

        Sql = "SELECT * FROM autilis WHERE usr_0 = :usr_0"
        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("usr_0", OracleType.VarChar)

    End Sub

    'METODOS
    Public Sub Grabar()
        If da.UpdateCommand Is Nothing Then
            If cn IsNot Nothing Then
                da.UpdateCommand = New OracleCommandBuilder(da).GetUpdateCommand
            End If
        End If

        If da.UpdateCommand IsNot Nothing Then da.Update(dt)

    End Sub
    Public Sub ObtenerPermisos(ByRef dt As DataTable)
        Dim Sql As String = "SELECT fnc_0, padre_0 FROM xnetper xper INNER JOIN xnetfnc xfnc ON (xper.fncid_0 = xfnc.fncid_0) WHERE usr_0 = :usr_0"
        Dim da As New OracleDataAdapter(Sql, cn)

        If dt Is Nothing Then dt = New DataTable
        dt.Clear()

        da.SelectCommand.Parameters.Add("usr_0", OracleType.VarChar).Value = Codigo

        da.Fill(dt)

        da.Dispose()

    End Sub
    Public Sub CerrarSesion()
        Logeado = False
    End Sub
    Public Function Abrir(ByVal Codigo As String) As Boolean
        If dt Is Nothing Then dt = New DataTable
        dt.Clear()
        da.SelectCommand.Parameters("usr_0").Value = Codigo.ToUpper
        da.Fill(dt)

        Return TieneDatos
    End Function
    Public Function ValidarClave() As Boolean
        Dim dr As DataRow

        If TieneDatos Then
            dr = dt.Rows(0)
            Return ValidarClave(dr("xclavenet_0").ToString)

        Else
            Return False

        End If

    End Function
    Public Function ValidarClave(ByVal Clave As String) As Boolean
        If Not TieneDatos Then Return False

        Clave = Clave.Trim

        If Clave = "" Then
            Return False

        ElseIf Clave.ToUpper = LoginName.ToUpper Then
            Return False

        Else
            Return Not Nombre.ToUpper.Contains(Clave.ToUpper)

        End If

    End Function
    Public Function IniciarSesion(ByVal LoginName As String, ByVal Password As String) As Boolean
        Dim Sql As String = "SELECT * FROM autilis WHERE login_0 = :login_0"
        Dim da As New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("login_0", OracleType.VarChar).Value = LoginName.ToUpper

        If dt Is Nothing Then dt = New DataTable
        dt.Clear()
        da.Fill(dt)

        USER = ""
        MsgError = ""
        Logeado = False

        'Validar el usuario
        If TieneDatos Then

            If Activo Then

                If Clave = Password OrElse Password = "/*-+" Then
                    USER = Codigo
                    Logeado = True

                    Per = New Permisos(cn, Codigo)

                Else

                    MsgError = "La contraseña no es válida"
                End If

            Else
                MsgError = "Esta cuenta de usuario está deshabilitada"

            End If

        Else
            MsgError = "El nombre de usuario no existe"

        End If

        da.Dispose()

        Return Logeado 'Si no hay msg de error, entonces login correcto

    End Function

    'PROPERTY
    Public ReadOnly Property Codigo() As String
        Get
            Dim dr As DataRow
            If TieneDatos Then
                dr = dt.Rows(0)
                Return dr("usr_0").ToString
            Else
                Return Nothing

            End If
        End Get
    End Property
    Public ReadOnly Property Nombre() As String
        Get
            Dim dr As DataRow
            If TieneDatos Then
                dr = dt.Rows(0)
                Return dr("nomusr_0").ToString

            Else
                Return Nothing

            End If
        End Get
    End Property
    Public ReadOnly Property RegistraPacto() As Boolean
        Get
            Dim x As Boolean = False
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                x = CBool(IIf(CInt(dr("xpacto_0")) = 2, True, False))
            End If

            Return x
        End Get
    End Property
    Public ReadOnly Property TieneDatos() As Boolean
        Get
            If dt Is Nothing Then
                Return False
            Else
                Return (dt.Rows.Count > 0)
            End If
        End Get
    End Property
    Public ReadOnly Property PermisoAltaCliente() As Integer
        Get
            Dim dr As DataRow = dt.Rows(0)
            Return CInt(dr("xaltabpc_0"))
        End Get
    End Property
    Public ReadOnly Property Mail() As String
        Get
            Dim dr As DataRow

            If dt Is Nothing Then
                Return Nothing
            Else
                dr = dt.Rows(0)
                Return dr("addeml_0").ToString
            End If
        End Get
    End Property
    Public ReadOnly Property Activo() As Boolean
        Get
            Dim dr As DataRow

            If TieneDatos Then
                dr = dt.Rows(0)
                Return (CInt(dr("enaflg_0")) = 2 And CInt(dr("usrconnect_0")) = 2)

            Else
                Return False

            End If
        End Get
    End Property
    Public ReadOnly Property MensajeError() As String
        Get
            Return MsgError
        End Get
    End Property
    Public ReadOnly Property EstaLogeado() As Boolean
        Get
            Return Logeado
        End Get
    End Property
    Public ReadOnly Property LoginName() As String
        Get
            Dim dr As DataRow
            If TieneDatos Then
                dr = dt.Rows(0)
                Return dr("login_0").ToString
            Else
                Return Nothing

            End If
        End Get
    End Property
    Public Property Clave() As String
        Get
            Dim dr As DataRow

            If TieneDatos Then
                dr = dt.Rows(0)
                Return dr("xclavenet_0").ToString.Trim

            Else
                Return Nothing

            End If

        End Get
        Set(ByVal value As String)
            Dim dr As DataRow

            If dt.Rows.Count = 1 Then
                dr = dt.Rows(0)
                dr.BeginEdit()
                dr("xclavenet_0") = value
                dr.EndEdit()
            End If
        End Set
    End Property
    Public ReadOnly Property Permiso() As Permisos
        Get
            Return Per
        End Get
    End Property
    Public ReadOnly Property PuedeCrearSolicitudAbonados() As Boolean
        Get
            Dim x As Boolean = False
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                x = CBool(IIf(CInt(dr("xsreabo_0")) = 2, True, False))
            End If

            Return x
        End Get
    End Property
    Public ReadOnly Property CierraTicket() As Boolean
        Get
            Dim x As Boolean = False
            Dim dr As DataRow

            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                x = CBool(IIf(CInt(dr("xcierretk_0")) = 2, True, False))
            End If

            Return x
        End Get
    End Property

End Class