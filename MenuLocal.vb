Imports System.Data.OracleClient
Imports System.Windows.Forms

Public Class MenuLocal
    Public dt As New DataTable
    Private cn As OracleConnection
    Private da As OracleDataAdapter
    Private MenuNro As Integer
    Private LineaCero As String = ""

    Public Sub New(ByVal cn As OracleConnection, Optional ByVal orden As Boolean = True)
        Dim Sql As String

        Sql = "SELECT * "
        Sql &= "FROM aplstd "
        Sql &= "WHERE lanchp_0 = :lanchp_0 AND lan_0 = 'SPA' AND lannum_0 >= 0 "
        If orden Then
            Sql &= "ORDER BY lannum_0"
        Else
            Sql &= "ORDER BY lanmes_0"
        End If
        Me.cn = cn

        da = New OracleDataAdapter(Sql, cn)
        da.SelectCommand.Parameters.Add("lanchp_0", OracleType.Number)
        da.InsertCommand = New OracleCommandBuilder(da).GetInsertCommand
        da.UpdateCommand = New OracleCommandBuilder(da).GetUpdateCommand

    End Sub
    Public Sub New(ByVal cn As OracleConnection, ByVal MenuNro As Integer, ByVal Blanco As Boolean, Optional ByVal orden As Boolean = True)

        Me.New(cn, orden)

        AbrirMenu(MenuNro, Blanco)

    End Sub
    Public Sub AbrirMenu(ByVal MenuNro As Integer, ByVal Blanco As Boolean)
        Dim dr As DataRow

        Me.MenuNro = MenuNro

        Try
            da.SelectCommand.Parameters("lanchp_0").Value = MenuNro

            dt.Clear()
            da.Fill(dt)

            ObtenerLineaCero()

            If Blanco Then
                dr = dt.NewRow
                dr("lanchp_0") = 0
                dr("lanmes_0") = " "
                dr("lannum_0") = 0
                dr("lan_0") = "SPA"
                dr("upddat_0") = #12/31/1599#
                dr("updusr_0") = " "
                dr("creusr_0") = " "
                dr("credat_0") = #12/31/1599#
                dr("cretim_0") = 0
                dr("updtim_0") = 0
                dt.Rows.InsertAt(dr, 0)
                dt.AcceptChanges()
            End If

        Catch ex As Exception

        End Try

    End Sub
    Private Sub ObtenerLineaCero()

        For Each dr As DataRow In dt.Rows

            If CInt(dr("lannum_0")) = 0 Then
                LineaCero = dr("lanmes_0").ToString
                dr.Delete()
                Exit For
            End If
        Next

        dt.AcceptChanges()

    End Sub
    Public Function Enlazar(ByVal cbo As DataGridViewComboBoxColumn) As DataGridViewComboBoxColumn
        With cbo
            .DataSource = dt
            .DisplayMember = "lanmes_0"
            .ValueMember = "lannum_0"
        End With

        Return cbo

    End Function
    Public Sub Enlazar(ByRef lst As ListBox)
        With lst
            .DataSource = dt
            .DisplayMember = "lanmes_0"
            .ValueMember = "lannum_0"
        End With
    End Sub
    Public Sub Enlazar(ByVal lst As CheckedListBox)
        With lst
            .DataSource = dt
            .DisplayMember = "lanmes_0"
            .ValueMember = "lannum_0"
        End With
    End Sub
    Public Sub Enlazar(ByRef cbo As ComboBox)
        With cbo
            .DataSource = dt
            .DisplayMember = "lanmes_0"
            .ValueMember = "lannum_0"
        End With
    End Sub
    Public ReadOnly Property Tabla1() As DataTable
        Get
            Return dt
        End Get
    End Property
    Public Function Agregar(ByVal Texto As String) As Integer
        Dim dr As DataRow
        Dim i As Integer = 0

        'Formateo el texto a agregar
        Texto = Texto.Trim.ToUpper

        'Busco que no exista en la tabla
        i = 0
        For Each dr In dt.Rows
            If dr("lanmes_0").ToString.Trim.ToUpper = Texto Then
                i = CInt(dr("lannum_0"))
                Return i
            End If
        Next

        'Busco el numero index más alto
        i = 0
        For Each dr In dt.Rows
            If CInt(dr("lannum_0")) > i Then i = CInt(dr("lannum_0"))
        Next

        i += 1

        dr = dt.NewRow
        dr("lanchp_0") = MenuNro
        dr("lanmes_0") = Texto.ToUpper
        dr("lannum_0") = i
        dr("lan_0") = "SPA"
        dr("upddat_0") = #12/31/1599#
        dr("updusr_0") = USER
        dr("creusr_0") = " "
        dr("credat_0") = #12/31/1599#
        dr("cretim_0") = 0
        dr("updtim_0") = 0
        dt.Rows.Add(dr)

        Return i

    End Function
    Public Sub Grabar()
        da.Update(dt)
    End Sub
    Public Sub ModificarTexto(ByVal Indice As Integer, ByVal Texto As String)
        Dim dr As DataRow

        For Each dr In dt.Rows
            If CInt(dr("lannum_0")) = Indice Then
                dr.BeginEdit()
                dr("lanmes_0") = Texto
                dr.EndEdit()
                Exit For
            End If
        Next
        dt.AcceptChanges()

    End Sub
    Public Sub EliminarCodigoIgual(ByVal Codigo As String)
        Dim dr As DataRow
        Dim s As String
        Dim x As Integer

        For i = dt.Rows.Count - 1 To 1 Step -1
            dr = dt.Rows(i) 'Obtengo la fila de la tabla de acompañantes

            x = CInt(dr("lannum_0"))

            'Elimino el registro si hay coincidencia de codigo
            s = LineaCero.Substring(x - 1, 1)
            If s = Codigo Then dr.Delete()
        Next

    End Sub
    Public ReadOnly Property Descripcion(ByVal id As Integer) As String
        Get
            Dim dr As DataRow
            Dim txt As String = ""

            For Each dr In dt.Rows
                If CInt(dr("lannum_0")) = id Then
                    txt = dr("lanmes_0").ToString
                    Exit For
                End If
            Next

            Return txt
        End Get
    End Property

End Class