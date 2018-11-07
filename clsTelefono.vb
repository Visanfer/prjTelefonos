Option Explicit On 
Imports MySql.Data.MySqlClient
Imports prjControl

Public Class clsTelefono
    Public mnEmpresa As Integer
    Public mnCodigo As Long
    Public msTipo As String
    Public msNumero As String
    Public msNombre As String
    Public mbPrivado As Boolean
    Public msExtension As String

    Public mbEsNuevo As Boolean
    Public Event evtBusTelefono()       ' evento desencadenado despues una busqueda

#Region " Funciones y Rutinas varias "

    Public Function mpsCodigo() As String
        mpsCodigo = "clsTelefono-" & mnEmpresa & "-" & mnCodigo
    End Function

    Public Sub New()
        mbEsNuevo = True
    End Sub

    Public Sub mrMandaEvento()
        RaiseEvent evtBusTelefono()
    End Sub

    Public Sub mrBuscaTelefono()
        'Dim loBusTelefono As New frmBusTelefonos
        'loBusTelefono.mrCargar(Me, mnEmpresa)
    End Sub

    Public Sub mrClonar(ByRef loTelefono As clsTelefono)
        loTelefono.mnEmpresa = mnEmpresa
        loTelefono.mnCodigo = mnCodigo
        loTelefono.msTipo = msTipo
        loTelefono.msNumero = msNumero
        loTelefono.msNombre = msNombre
        loTelefono.msExtension = msExtension
        loTelefono.mbEsNuevo = mbEsNuevo
    End Sub

#End Region

#Region " Acceso a la base de Datos "

    Public Sub mrRecuperaDatos()
        Dim lconConexion As mySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim lsSql As String
        Dim loRecord As mySqlDataReader

        lsSql = "select * from telefonos where emp_tel = " & mnEmpresa & _
                " and cod_tel = " & mnCodigo & _
                " and tip_tel = '" & msTipo & "'"
        Dim loComando As New mySqlCommand(lsSql, lconConexion)
        loRecord = loComando.ExecuteReader
        mbEsNuevo = True
        While loRecord.Read
            mnEmpresa = mfnInteger(loRecord("emp_tel") & "")
            mnCodigo = mfnLong(loRecord("cod_tel") & "")
            msTipo = Trim(loRecord("tip_tel") & "")
            msNumero = Trim(loRecord("num_tel") & "")
            msNombre = Trim(loRecord("nom_tel") & "")
            mbPrivado = (InStr(msNombre, "PRIVADO") > 0)
            msExtension = Trim(loRecord("ext_tel") & "")
            mbEsNuevo = False
        End While
        loRecord.Close()
        lconConexion.Close()

    End Sub

    Public Sub mrGrabaDatos()
        Dim lconConexion As mySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim lsSql As String
        Dim loComando As New mySqlCommand

        If mbEsNuevo Then
            lsSql = "insert into telefonos values ('" & mnEmpresa & "','" & _
                    mnCodigo & "','" & _
                    msTipo & "','" & _
                    msNumero & "','" & _
                    msNombre & "','" & _
                    msExtension & "')"
            loComando = New mySqlCommand(lsSql, lconConexion)
            loComando.ExecuteNonQuery()
            lconConexion.Close()
        Else
            lsSql = "update telefonos set num_tel = '" & msNumero & _
                "', nom_tel = '" & msNombre & _
                "', ext_tel = '" & msExtension & _
                "' where emp_tel = " & mnEmpresa & _
                " and cod_tel = " & mnCodigo & _
                " and tip_tel = '" & msTipo & "'"
            loComando = New mySqlCommand(lsSql, lconConexion)
            loComando.ExecuteNonQuery()
            lconConexion.Close()
        End If
        mbEsNuevo = False

    End Sub

    Public Sub mrBorrarDatos()
        Dim lconConexion As mySqlConnection = mfconConexionSQL(False)
        If lconConexion.State = ConnectionState.Closed Then Exit Sub

        Dim lsSql As String
        Dim loComando As New mySqlCommand

        ' ahora borro de del fichero de cabeceras ****************
        lsSql = "delete from telefonos where emp_tel = " & mnEmpresa & _
                " and cod_tel = " & mnCodigo & _
                " and tip_tel = '" & msTipo & "'"
        loComando = New mySqlCommand(lsSql, lconConexion)
        loComando.ExecuteNonQuery()
        lconConexion.Close()
        mbEsNuevo = True

    End Sub

#End Region

End Class
