Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Module ConexionSQL
    Public ConexionSQL As SqlConnection

    Public servidor As String = "NDCA0WPAP138.MX.CKSNNET.COM\ALPHAMTY"
    Public bd As String = "OEE_data"
    Public usuario As String = "SA"
    Public pass As String = "Alpha123"

    Public conexion As New OleDbConnection
    Public comando As New SqlCommand
    Public adaptador As New SqlDataAdapter
    Public registro As New DataSet

    Public estadosql As String

    Sub Conexion_SQL()
        Try
            Dim cadena = "data source =" & servidor & "; initial catalog = " & bd & "; user id = " & usuario & "; password = " & pass
            'Dim cadena_OEEdata = "data source =" & servidor_OEEdata & "; initial catalog =" & bd_OEEdata & "; Integrated Security = SSPI"

            ConexionSQL = New SqlConnection(cadena)

            ConexionSQL.Open()
            estadosql = "Online_OEEdata"
            'MsgBox("Conexión establecida")
            Form1.Label1.Visible = True
            Form1.Label2.Visible = False

            'ConexionSQL_OEEdata.Close()
            'estadosql = "Offline"
            'MsgBox("Conexión cerrada")

        Catch ex As Exception

            estadosql = "Failure offline"
            Form1.Label1.Visible = False
            Form1.Label2.Visible = True
            'MsgBox("Error al tratar de conectar a la base de datos: " & ex.Message)

        End Try
    End Sub

End Module
