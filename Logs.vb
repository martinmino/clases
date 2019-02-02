Imports System.IO

Public Class Logs
    Private Archivo As String = ""
    Private sw As StreamWriter

    Public Sub Nuevo()
        'Genero el nombre del archivo
        Archivo = "logs\Intervenciones_Automaticas_" & Now.ToString("yyyy-MM-dd-HH-mm") & ".log"

        Try
            If Not File.Exists("logs") Then
                Directory.CreateDirectory("logs")
            End If
            sw = New StreamWriter(Archivo)
        Catch ex As Exception
        End Try

    End Sub
    Public Sub Escribir(Optional ByVal txt As String = "")
        sw.WriteLine(txt)
    End Sub
    Public Sub Cerrar()
        sw.Close()
        sw.Dispose()
        sw = Nothing
    End Sub

End Class