Imports System.Data.OracleClient

Public Module Module1

    Friend USER As String = "ADM"

    Friend Const ARTICULO_PRESTAMO_EXT As String = "601003"
    Friend Const ARTICULO_PRESTAMO_MAN As String = "607003"
    Friend Const ARTICULO_FLETE_Y_ACARREO As String = "653005"
    Public Const DB_USR As String = "GEOPROD"
    Public Const DB_SRV As String = "adx"
    Friend Const RPTX3 As String = "\\" & DB_SRV & "\Folders\" & DB_USR & "\REPORT\SPA\"

    Friend Sub Parametro(ByVal cm As OracleCommand, ByVal Nombre As String, ByVal Tipo As OracleType, ByVal Version As DataRowVersion, Optional ByVal Campo As String = "", Optional ByVal Valor As Object = Nothing)
        With cm.Parameters
            With .Add(Nombre, Tipo)
                If Valor Is Nothing Then
                    .SourceVersion = Version
                    If Campo = "" Then
                        .SourceColumn = .ParameterName
                    Else
                        .SourceColumn = Campo
                    End If

                Else
                    .Value = Valor

                End If

            End With
        End With

    End Sub


End Module