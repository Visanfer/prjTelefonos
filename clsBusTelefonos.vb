Option Explicit On 
Imports MySql.Data.MySqlClient
Imports prjControl

Public Class clsBusTelefonos
    Public mnEmpresa As Integer
    Public mnCodigo As Integer
    Public msNumero As String
    Public msNombre As String

    Public mcolTelefonos As Collection

    Public Sub mrBuscaTelefonos(ByVal lsTipo As String)
        Dim lconConexion As MySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim loTelefono As clsTelefono
        Dim loRecord As MySqlDataReader
        Dim lsSql As String

        mcolTelefonos = New Collection
        lsSql = "select * from telefonos where tip_tel = '" & lsTipo & "' order by cod_tel asc"
        Dim loComando As New MySqlCommand(lsSql, lconConexion)
        loRecord = loComando.ExecuteReader
        While loRecord.Read
            loTelefono = New clsTelefono
            loTelefono.mnEmpresa = mfnInteger(loRecord("emp_tel") & "")
            loTelefono.mnCodigo = mfnLong(loRecord("cod_tel") & "")
            loTelefono.msTipo = Trim(loRecord("tip_tel") & "")
            loTelefono.msNumero = Trim(loRecord("num_tel") & "")
            loTelefono.msNombre = Trim(loRecord("nom_tel") & "")
            loTelefono.mbPrivado = (InStr(loTelefono.msNombre, "PRIVADO") > 0)
            loTelefono.msExtension = Trim(loRecord("ext_tel") & "")
            loTelefono.mbEsNuevo = False
            mcolTelefonos.Add(loTelefono, loTelefono.mpsCodigo)
        End While
        loRecord.Close()
        lconConexion.Close()

    End Sub

End Class
