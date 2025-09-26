'##############################################################
' DbConnectionManager.vb
' Purpose: Centralized ADO.NET connection factory for SQL Server.
' Dependencies: System.Configuration, System.Data.SqlClient
' Current date: 2025-09-25
'##############################################################
Imports System.Configuration
Imports System.Data.SqlClient

Public NotInheritable Class DbConnectionManager
    Private Sub New()
    End Sub

    ''' <summary>
    ''' Creates and opens a SQL Server connection using the "GigGearCoversDb" connection string.
    ''' </summary>
    ''' <returns>An open SqlConnection instance.</returns>
    ''' <exception cref="InvalidOperationException">Thrown if the connection string is missing.</exception>
    ''' <exception cref="SqlException">Thrown if the connection cannot be opened.</exception>
    Public Shared Function CreateOpenConnection() As SqlConnection
        Dim cs As String = ConfigurationManager.ConnectionStrings("GigGearCoversDb")?.ConnectionString
        If String.IsNullOrWhiteSpace(cs) Then
            Throw New InvalidOperationException("Connection string 'GigGearCoversDb' is missing in App.config.")
        End If
        Try
            Dim conn As New SqlConnection(cs)
            conn.Open()
            Return conn
        Catch ex As SqlException
            Throw New InvalidOperationException("Failed to open SQL connection: " & ex.Message, ex)
        End Try
    End Function

End Class
