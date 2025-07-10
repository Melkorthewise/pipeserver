Imports System.IO
Imports System.IO.Pipes
Imports System.Text
Imports System.Threading.Tasks

Public Class PipeClientHelper

    Private Const PipeName As String = "TestPipe"

    Public Shared Async Function SendMessageAsync(message As String, Optional timeoutMs As Integer = 3000) As Task(Of String)
        Using pipeClient As New NamedPipeClientStream(".", PipeName, PipeDirection.InOut, PipeOptions.Asynchronous)
            Try
                Await pipeClient.ConnectAsync(timeoutMs)
            Catch ex As Exception
                Throw New IOException("Failed to connect to the pipe server: " & ex.Message, ex)
            End Try

            ' Send message
            Dim sendBytes = Encoding.UTF8.GetBytes(message)
            Await pipeClient.WriteAsync(sendBytes, 0, sendBytes.Length)
            Await pipeClient.FlushAsync()

            ' Receive response
            Dim buffer(255) As Byte
            Dim bytesRead = Await pipeClient.ReadAsync(buffer, 0, buffer.Length)
            If bytesRead > 0 Then
                Return Encoding.UTF8.GetString(buffer, 0, bytesRead)
            Else
                Return String.Empty
            End If
        End Using
    End Function

End Class
