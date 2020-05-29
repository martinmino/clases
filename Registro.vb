Imports Microsoft.Win32

Public Class Registro
    Shared Function getDato(ByVal Clave As String) As String
        Dim RegKey As RegistryKey
        Dim s As String = ""

        Try
            'Creo o abro clave del registro
            RegKey = Registry.CurrentUser.OpenSubKey("Software", True).CreateSubKey("Georgia")
            s = RegKey.GetValue(Clave).ToString
            RegKey.Close()

        Catch ex As Exception
        End Try

        Return s

    End Function
    Shared Sub setDato(ByVal Clave As String, ByVal Valor As String)
        Dim RegKey As RegistryKey

        Try
            'Creo o abro clave del registro
            RegKey = Registry.CurrentUser.OpenSubKey("Software", True).CreateSubKey("Georgia")
            RegKey.SetValue(Clave, Valor)
            RegKey.Close()

        Catch ex As Exception
        End Try

    End Sub

End Class