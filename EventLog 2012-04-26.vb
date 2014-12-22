Option Explicit On
Option Strict On
Imports System
Imports System.Diagnostics
Imports System.Threading

Public Class EventLog

    'Public Shared Function Write(ByVal AppName As String, ByVal ErrorMessage As String, ByVal ErrType As Integer)
    '
    '   Dim appLog As New System.Diagnostics.EventLog
    '      appLog.Source = AppName
    '
    '   If ErrType = 1 Then
    '      appLog.WriteEntry(ErrorMessage, EventLogEntryType.Warning)
    '   ElseIf ErrType = 2 Then
    '      appLog.WriteEntry(ErrorMessage, EventLogEntryType.Error)
    '   Else 0
    '        appLog.WriteEntry(ErrorMessage, EventLogEntryType.Information)
    '   End If
    '   End Function

    Public Shared Function WriteToEventLog(ByVal Entry As String, ByVal eltType As EventLogEntryType) As Boolean
        Dim appName As String = "CIM_PreVersion"
        Dim logName As String = "App CIM PreVersion"

        Try
            'Register the App as an Event Source
            If Not System.Diagnostics.EventLog.SourceExists(appName) Then
                System.Diagnostics.EventLog.CreateEventSource(appName, logName)
            End If

            Dim myEventLog As New System.Diagnostics.EventLog()
            myEventLog.Source = appName
            'WriteEntry is overloaded; this is one of 10 ways to call it
            myEventLog.WriteEntry(Entry, eltType)
            Return True

        Catch Ex As Exception
            'objEventLog.WriteEntry("CIM Service - Function: WriteToEventLog " & vbCrLf & "Exception: " & Ex.Message, EventLogEntryType.Error)
            Console.WriteLine("CIM Service Exception: Function WriteToEventLog" & vbCrLf & "Exception Message: " & Ex.Message.ToString() & vbCrLf & "Exception Target: " & Ex.TargetSite.ToString())
            Return False
        End Try
    End Function
End Class
                          