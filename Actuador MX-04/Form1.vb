Imports System.Data.SqlClient
Imports System.IO

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Conexion_SQL()
        conexion.Close()

        Me.WindowState = FormWindowState.Minimized

    End Sub

    Private Sub T_Status_Tick(sender As Object, e As EventArgs) Handles T_Status.Tick


        Conexion_SQL()
        conexion.Close()
        CheckIfisOpen()
        CheckIfRunning()


        'Generar records en una tabla de cierta base de datos

        'Obtenemos la Hora y Fecha
        Dim Fecha_log As String = DateTime.Now.ToString("MM/dd/yyyy")
        Dim Hora_log As String = Date.Now.ToString("HH:mm:ss")

        Dim EstadoPC As String

        If Label1.Visible = True Then
            EstadoPC = "Online"
        Else
            EstadoPC = "Offline"
        End If

        'Capturamos si la ventana del Indusoft esta abierta
        Dim VentanaIduSoft As String

        If Label5.Visible = True Then
            VentanaIduSoft = "Open"
        Else
            VentanaIduSoft = "Closed"
        End If

        'Capturamos si la aplicación del Indusoft esta corriendo
        Dim StatusIduSoft As String

        If Label4.Visible = True Then
            StatusIduSoft = "Running"
        Else
            StatusIduSoft = "Stoped"
        End If

        'Generamos el Log con la cadena de SQL
        Try
            Dim cadena_save As String

            cadena_save = "INSERT INTO [Fluxmonitor].[dbo].[Log_MX04] ([Fecha]
      ,[Hora]
      ,[Estado_PC]
      ,[Ventana_InduSoft]
      ,[Status_Aplicacion]) VALUES ('" & Fecha_log & "'
      ,'" & Hora_log & "'
      ,'" & EstadoPC & "'
      ,'" & VentanaIduSoft & "'
      ,'" & StatusIduSoft & "')"

            Dim comando_save As SqlCommand
            comando_save = New SqlCommand(cadena_save, ConexionSQL.ConexionSQL)
            comando_save.ExecuteNonQuery()

        Catch ex As Exception
            'MsgBox("PROBLEMA CON DISPARAR DATOS AL RESUMEN:" & ex.Message)
        End Try


    End Sub

    Dim p() As Process

    Private Sub CheckIfRunning()
        p = Process.GetProcessesByName("Viewer")
        If p.Count > 0 Then
            ' Process is running
            'MsgBox("Esta corriendo la aplicación")
            Label3.Visible = False
            Label4.Visible = True
        Else
            ' Process is not running
            'MsgBox("No esta corriendo la aplicación")
            Label3.Visible = True
            Label4.Visible = False
        End If
    End Sub

    Dim s() As Process

    Private Sub CheckIfisOpen()
        s = Process.GetProcessesByName("Studio Manager")
        If s.Count > 0 Then
            'Process is running
            'MsgBox("Esta abierta la aplicación")
            Label6.Visible = False
            Label5.Visible = True
        Else
            'Process is not running
            'MsgBox("No esta abierta la aplicación")
            Label6.Visible = True
            Label5.Visible = False
        End If
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        e.Cancel = True
    End Sub
End Class
