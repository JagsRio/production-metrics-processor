' ===============================================================================
' Production Metrics Reporter Windows Service
' Language: VB.NET (.NET Framework)
' Purpose:  Monitors system metrics and sends automated reports
' Author:   [Your Name]
' GitHub:   https://github.com/[your-username]
' ===============================================================================
'
' Demonstrates:
'   • Windows Services with scheduled timers
'   • Database querying and data aggregation
'   • Report generation and email alerting
'   • Configuration management
'   • Robust error handling
'
' Sanitized for portfolio: Removed proprietary schemas and identifiers.
' ===============================================================================

Imports System.Timers
Imports System.IO
Imports System.Net.Mail
Imports System.Data.SqlClient
Imports System.Text

Public Class MetricsReporter
    Inherits System.ServiceProcess.ServiceBase

    Private WithEvents _reportTimer As New Timer()
    Private _logPath As String = "C:\Logs\MetricsReporter.log" ' Configurable

    Protected Overrides Sub OnStart(ByVal args() As String)
        Log("Service started.")

        ' Interval: 1 hour (configurable)
        _reportTimer.Interval = Double.Parse(ConfigurationManager.AppSettings("MonitoringIntervalMs") Or "3600000")
        _reportTimer.Start()
    End Sub

    Protected Overrides Sub OnStop()
        _reportTimer.Stop()
        Log("Service stopped.")
    End Sub

    Private Sub OnTimerElapsed(sender As Object, e As ElapsedEventArgs) Handles _reportTimer.Elapsed
        Dim currentHour = DateTime.Now.Hour
        If currentHour = 5 Then ' Configurable daily trigger
            Log("Starting daily metrics report generation.")
            GenerateAndSendReports()
        End If
    End Sub

    Private Sub GenerateAndSendReports()
        Try
            Dim clientMetrics = FetchClientMetrics()
            Dim volumeReport = BuildVolumeReport(clientMetrics)
            Dim staleReport = BuildStaleActivityReport(clientMetrics)
            Dim performanceReport = BuildPerformanceExceptionReport(clientMetrics)

            SendEmailReport("Daily Volume Summary", volumeReport)
            SendEmailReport("Stale Client Activity", staleReport)
            SendEmailReport("Performance Exceptions", performanceReport)
        Catch ex As Exception
            Log($"Error generating reports: {ex.Message}")
        End Try
    End Sub

    Private Function FetchClientMetrics() As DataTable
        ' Generalized query for client activity and performance
        Return QueryDatabase(
            "SELECT ClientID, Account, OfficeName, LastActivityDate, " &
            "SystemName, SystemVersion " &
            "FROM ClientMetricsView " &
            "ORDER BY OfficeName")
    End Function

    Private Function BuildVolumeReport(data As DataTable) As String
        ' Aggregates total processing volume
        Dim sb As New StringBuilder("Daily Processing Volume Report" & vbCrLf)
        ' Logic to count and summarize - omitted for brevity
        Return sb.ToString()
    End Function

    Private Function BuildStaleActivityReport(data As DataTable) As String
        Dim sb As New StringBuilder("Stale Activity Report" & vbCrLf)
        sb.AppendLine("ClientID\tAccount\tLastDate\tSystem\tVersion\tOffice")

        For Each row As DataRow In data.Rows
            If CDate(row("LastActivityDate")) < DateTime.Today.AddDays(-2) Then
                sb.AppendLine($"{row("ClientID")}\t{row("Account")}\t" &
                              $"{row("LastActivityDate")}\t{row("SystemName")}\t" &
                              $"{row("SystemVersion")}\t{row("OfficeName")}")
            End If
        Next

        Return sb.ToString()
    End Function

    Private Function BuildPerformanceExceptionReport(data As DataTable) As String
        Dim sb As New StringBuilder("Performance Exceptions Report" & vbCrLf)
        sb.AppendLine("ClientID\tAccount\tSuccessRate\tOffice")

        For Each row As DataRow In data.Rows
            Dim successRate = CalculateSuccessRate(row("ClientID"))
            If successRate < 0.5 Then ' 50% threshold
                sb.AppendLine($"{row("ClientID")}\t{row("Account")}\t" &
                              $"{successRate:P}\t{row("OfficeName")}")
            End If
        Next

        Return sb.ToString()
    End Function

    Private Function CalculateSuccessRate(clientId As String) As Double
        ' Placeholder for aggregated success metric calculation
        Return 0.75 ' Example value
    End Function

    Private Sub SendEmailReport(subject As String, body As String)
        ' In production: Uses SmtpClient with configured server/credentials
        Log($"Report sent: {subject} ({body.Split(vbCrLf).Length} lines)")
    End Sub

    Private Function QueryDatabase(query As String) As DataTable
        Dim dt As New DataTable()
        Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("MetricsDB").ConnectionString),
              adapter As New SqlDataAdapter(query, conn)
            adapter.Fill(dt)
        End Using
        Return dt
    End Function

    Private Sub Log(message As String)
        File.AppendAllText(_logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {message}{vbCrLf}")
    End Sub

End Class

Public Class MetricsSummary
    ' Generalized model for metrics data
    Public Property ClientId As String
    Public Property TotalProcessed As Integer
    Public Property SuccessCount As Integer
    ' Additional properties as needed
End Class