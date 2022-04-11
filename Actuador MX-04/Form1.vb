Imports System.Data.SqlClient
Imports System.IO
Imports System.Net.Mail

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Conexion_SQL()
        conexion.Close()

        Me.WindowState = FormWindowState.Minimized

    End Sub

    Private Sub T_Status_Tick(sender As Object, e As EventArgs) Handles T_Status.Tick

        Conexion_SQL()
        conexion.Close()
        Checkdelay()
        'CheckIfRunning()

        'Registra el estado de salud del robot

        'Obtenemos la Hora y Fecha
        Dim Fecha_log As String = DateTime.Now.ToString("MM/dd/yyyy")
        Dim Hora_log As String = Date.Now.ToString("HH:mm:ss")

        'Generamos el Log con la cadena de SQL
        Try
            Dim cadena_save As String

            cadena_save = "UPDATE [AlteaDB].[dbo].[altea_bot] " &
            "SET [bot_date] = '" & Fecha_log & "',
             [bot_hour]= '" & Hora_log & "'
             WHERE control = 1 "

            Dim comando_save As SqlCommand
            comando_save = New SqlCommand(cadena_save, ConexionSQL.ConexionSQL)
            comando_save.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox("PROBLEMA CON DISPARAR DATOS AL RESUMEN:" & ex.Message)
        End Try


    End Sub

    Dim p() As Process

    Private Sub CheckIfRunning()
        p = Process.GetProcessesByName("Viewer")
        If p.Count > 0 Then
            ' Process is running
            'MsgBox("Esta corriendo la aplicación")
            'Label3.Visible = False
            'Label4.Visible = True
        Else
            ' Process is not running
            'MsgBox("No esta corriendo la aplicación")
            'Label3.Visible = True
            'Label4.Visible = False
        End If
    End Sub


    Private Sub Checkdelay()

        Try
            'Revisa si hay registros con mas de 24 Horas


            Dim cadena_check As String
            Dim listado As String

            cadena_check = "SELECT top 1 batch, analyst, status, approve_date, approve_hour, DATEDIFF(HOUR , CONVERT(DATETIME, approve_date) + CONVERT(DATETIME, approve_hour), GETDATE()) as [hour_lapse] FROM [AlteaDB].[dbo].[batch_route] WHERE status = 0"

            Dim comando_check As SqlCommand
            comando_check = New SqlCommand(cadena_check, ConexionSQL.ConexionSQL)
            comando_check.ExecuteNonQuery()

            adaptador = New SqlDataAdapter(cadena_check, ConexionSQL.ConexionSQL)
            ConexionSQL.registro = New DataSet
            adaptador.Fill(registro, "Tabla3")
            listado = registro.Tables("Tabla3").Rows.Count

            Dim hora_lapso As Integer
            Dim lote As String
            hora_lapso = registro.Tables("Tabla3").Rows(0).Item("hour_lapse")
            lote = registro.Tables("Tabla3").Rows(0).Item("batch")

            If hora_lapso > 23 Then


                'Enviar Correo 

                Dim Smtp_Server As New SmtpClient
                Dim e_mail As New MailMessage
                Smtp_Server.UseDefaultCredentials = False
                'Smtp_Server.Credentials = New Net.NetworkCredential("DoNotReply@ElementSolutionsInc.com", "")
                'Smtp_Server.Port = 587
                'Smtp_Server.EnableSsl = True
                Smtp_Server.Host = "NDCA0SMTP01.pahprod.pahglobal.net"


                Dim cadena_correos As String
                Dim lista_correos As String

                cadena_correos = "SELECT [config1] ,[correos] FROM [AlteaDB].[dbo].[Tabla3] WHERE config1 = 'tres'"

                Dim comando_mails As SqlCommand
                comando_mails = New SqlCommand(cadena_correos, ConexionSQL.ConexionSQL)
                comando_mails.ExecuteNonQuery()

                adaptador = New SqlDataAdapter(cadena_correos, ConexionSQL.ConexionSQL)
                ConexionSQL.registro = New DataSet
                adaptador.Fill(registro, "Tabla3")
                lista_correos = registro.Tables("Tabla3").Rows.Count

                Dim Destinatarios As String
                Destinatarios = registro.Tables("Tabla3").Rows(0).Item("correos")


                'Dim Destinatarios As String
                'Destinatarios = registro.Tables("Tabla3").Rows(0).Item("correos")

                'MsgBox(Destinatarios)

                e_mail = New MailMessage
                e_mail.From = New MailAddress("notification.Alphamty@macdermidalpha.com")
                e_mail.To.Add(Destinatarios.ToString) '<--correos destino
                e_mail.Subject = ("ALTEA: -Alerta de pasta con mas de 24 Horas")
                e_mail.IsBodyHtml = False
                e_mail.Body = ("ALERTA, La Pasta con Lote: " & lote & ", Excede las 24 horas desde su liberación sin haber sido ingresada al refrigerador")

                Smtp_Server.Send(e_mail)
                'MsgBox("Correo enviado")
                'btnEnviar.Enabled = True
                'txtAsunto.Text = ""
                'txtMensaje.Text = ""


                'ACTUALIZAR EL ESTATUS QUE YA SE ENVIO NOTIFICACION

                Dim cadena_actualiza As String

                cadena_actualiza = "UPDATE [AlteaDB].[dbo].[batch_route] " &
                "SET [status] = 1 
                WHERE [batch] = '" & lote & "'"

                Dim comando_actualiza As SqlCommand
                comando_actualiza = New SqlCommand(cadena_actualiza, ConexionSQL.ConexionSQL)
                comando_actualiza.ExecuteNonQuery()

            End If

        Catch ex As Exception
            'MsgBox("Problema con inspeccionar base de datos: " & ex.Message)
        End Try

    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        e.Cancel = True
    End Sub
End Class
