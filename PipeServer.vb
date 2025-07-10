Imports System.IO.Pipes
Imports System.Security.AccessControl
Imports System.Security.Principal
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks

Public Class PipeServerHandler

    Private ReadOnly PipeName As String = "TestPipe"
    Private ReadOnly Log As Action(Of String)

    Public Sub New(logAction As Action(Of String))
        Log = logAction
    End Sub

    Public Async Sub Run(ct As CancellationToken)
        While Not ct.IsCancellationRequested
            Using pipeServer As New NamedPipeServerStream(PipeName, PipeDirection.InOut, 1,
                                              PipeTransmissionMode.Message,
                                              PipeOptions.Asynchronous,
                                              0, 0,
                                              CreatePipeSecurity())
                Log("Waiting for client connection...")
                Try
                    Await pipeServer.WaitForConnectionAsync(ct)
                Catch ex As OperationCanceledException
                    Log("Pipe server canceled.")
                    Exit While
                Catch ex As Exception
                    Log("WaitForConnection error: " & ex.Message)
                    Continue While
                End Try

                Log("Client connected.")

                Try
                    Dim buffer(255) As Byte
                    Dim bytesRead = Await pipeServer.ReadAsync(buffer, 0, buffer.Length, ct)
                    If bytesRead > 0 Then
                        Dim receivedMsg = Encoding.UTF8.GetString(buffer, 0, bytesRead)
                        Log("Received from client: " & receivedMsg)

                        ' Respond with a hardcoded string
                        Dim responseMsg = "Hello from service"
                        Dim responseBytes = Encoding.UTF8.GetBytes(responseMsg)
                        Await pipeServer.WriteAsync(responseBytes, 0, responseBytes.Length, ct)
                        Await pipeServer.FlushAsync(ct)
                        Log("Sent response to client.")
                    End If
                Catch ex As Exception
                    Log("Communication error: " & ex.Message)
                End Try

                pipeServer.Disconnect()
            End Using
        End While
        Log("Pipe server stopped.")
    End Sub

    Private Function CreatePipeSecurity() As PipeSecurity
        Dim pipeSecurity = New PipeSecurity()
        pipeSecurity.AddAccessRule(New PipeAccessRule(New SecurityIdentifier(WellKnownSidType.WorldSid, Nothing),
                                                  PipeAccessRights.FullControl,
                                                  AccessControlType.Allow))
        Return pipeSecurity
    End Function

End Class
