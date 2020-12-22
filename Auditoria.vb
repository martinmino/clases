Imports System.Data.OracleClient

Public Class Auditoria
    Private cn As OracleConnection
    Private cm As OracleCommand

    Public Sub New(ByVal cn As OracleConnection)
        Dim sql As String

        Me.cn = cn

        sql = "insert into aespion values (:fecha, :hora, :usuario, :funcion, :tabla, :accion, :id, :id2) "

        cm = New OracleCommand(Sql, cn)
        cm.Parameters.Add("fecha", OracleType.DateTime)
        cm.Parameters.Add("hora", OracleType.VarChar)
        cm.Parameters.Add("usuario", OracleType.VarChar)
        cm.Parameters.Add("funcion", OracleType.VarChar)
        cm.Parameters.Add("tabla", OracleType.VarChar)
        cm.Parameters.Add("accion", OracleType.Number)
        cm.Parameters.Add("id", OracleType.VarChar)
        cm.Parameters.Add("id2", OracleType.VarChar)

    End Sub
    Public Sub Grabar(ByVal usuario As String, ByVal funcion As String, ByVal tabla As String, ByVal accion As Integer, ByVal id As String)

        cm.Parameters("fecha").Value = Today
        cm.Parameters("hora").Value = Now.ToString("HH:mm:ss")
        cm.Parameters("usuario").Value = usuario
        cm.Parameters("funcion").Value = funcion
        cm.Parameters("tabla").Value = tabla
        cm.Parameters("accion").Value = accion
        cm.Parameters("id").Value = id
        cm.Parameters("id2").Value = " "

        Try
            If cn.State <> ConnectionState.Open Then cn.Open()
            cm.ExecuteNonQuery()

        Catch ex As Exception

        End Try

    End Sub

End Class
