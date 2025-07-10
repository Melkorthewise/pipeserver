Imports System.ServiceProcess
Imports System.Threading

Public Class DSservice
    Inherits ServiceBase

    Private _cts As CancellationTokenSource
    Private server As PipeServerHandler

    Public Sub New()
        Me.ServiceName = "Service"
    End Sub

    Protected Overrides Sub OnStart(args() As String)
        _cts = New CancellationTokenSource()
        server = New PipeServerHandler(AddressOf Log)
        Task.Run(Sub() server.Run(_cts.Token))
        Log("Service started, pipe server running.")
    End Sub

    Protected Overrides Sub OnStop()
        If _cts IsNot Nothing Then
            _cts.Cancel()
            Log("Service stopping.")
        End If
    End Sub
End Class

Public Module Logger
    Private ReadOnly SourceName As String = "ServiceSource"
    Private ReadOnly LogName As String = "Application"

    Public Sub Log(message As String, Optional entryType As EventLogEntryType = EventLogEntryType.Information)
        Try
            ' Register once (optional: remove this after installer handles it)
            If Not EventLog.SourceExists(SourceName) Then
                EventLog.CreateEventSource(SourceName, LogName)
            End If

            EventLog.WriteEntry(SourceName, message, entryType)
        Catch ex As Exception
            ' Fallback (optional): write to a file or ignore
        End Try
    End Sub
End Module