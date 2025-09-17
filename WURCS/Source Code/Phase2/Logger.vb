Imports System.IO
Imports System.Threading

Module Logger
    Private ReadOnly logFilePath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "application.log")
    Private ReadOnly logLock As New Object()

    ''' <summary>
    ''' Logs a message to the log file with a timestamp.
    ''' </summary>
    ''' <param name="message">The message to log.</param>
    Public Sub Log(message As String)
        Dim logEntry As String = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} - {message}"
        SyncLock logLock
            File.AppendAllText(logFilePath, logEntry & Environment.NewLine)
        End SyncLock
    End Sub
End Module
