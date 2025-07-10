Imports System.IO.Pipes
Imports System.Text

Public Class Form1

    Private Async Sub btnSend_Click(sender As Object, e As EventArgs) Handles btnSend.Click
        Try
            Dim response = Await PipeClientHelper.SendMessageAsync("Hello from client")
            MessageBox.Show("Response: " & response)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub

End Class
