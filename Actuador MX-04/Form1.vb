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
