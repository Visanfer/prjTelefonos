Option Explicit On
Imports prjControl
Imports MySql.Data.MySqlClient

Public Module modTelefonos
    Public gnLlave As Integer          ' llave actual de seguridad
    Public goUsuario As New clsUsuario
    Public goProfile As New clsProfileLocal

    Public Sub Main()
        Dim loPrincipal As New frmTelefonos
        loPrincipal.mrCargar(1, Nothing)
    End Sub

End Module
