Imports System.IO
Imports System.Data.OracleClient

Public Class EtiquetaService
    Private Const ETIQUETA As String = "PLANTILLAS\etiqueta_logistica.txt"

    Private itn As Intervencion
    Private cn As OracleConnection

    Public Sub New(ByVal cn As OracleConnection)
        Me.cn = cn
    End Sub
    Public Sub Abrir(ByVal itn As Intervencion)
        Me.itn = itn
    End Sub
    Public Sub Abrir(ByVal NumeroIntervencion As String)
        If Me.itn Is Nothing Then itn = New Intervencion(cn)
        itn.Abrir(NumeroIntervencion)
    End Sub
    Private Function ArchivoSalida() As String
        Return "Etiqueta_" & itn.Numero & ".txt"
    End Function
    Public Sub Imprimir()

        Dim txt As String = ""
        Dim sr As StreamReader
        Dim sw As StreamWriter
        Dim Copias As Integer = 0
        Dim Numero As String = ""

        Try
            'Abro plantilla de etiqueta
            sr = New StreamReader(ETIQUETA)
            txt = sr.ReadToEnd
            sr.Close()

            'Calculo cantidad de copias
            Copias = itn.CantidadEquiposTeoricos \ 2 + 1

            'Obtengo número de intervención
            Numero = itn.Numero.Substring(5)

            'Incluyo los datos en la plantailla
            txt = txt.Replace("{itn}", Numero)
            txt = txt.Replace("{cantidad}", Copias.ToString)

            'Guardo la etiqueta en el archivo de salida
            sw = New StreamWriter(ArchivoSalida)
            sw.Write(txt)
            sw.Close()

            'Envio archivo de salida a la impresora
            Dim prn As New Impresora(cn, "LOGISTICA")

            File.Copy(ArchivoSalida, "\\mmino-pc\zebra") 'prn.RecursoRed)

        Catch ex As Exception

        End Try
    End Sub
End Class