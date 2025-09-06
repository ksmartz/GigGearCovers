Imports System.Text.Json
Imports System.Windows.Forms
Imports System
Imports System.IO


Public Class frmDashboard

    ' ---- Update these constants as needed ----
    Private Const SchemaTaskName As String = "export_vbGGC_DailySchemaToJson"
    Private ReadOnly SchemaJsonPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Scripts\Schema\DailySchema.json")
    Private Const WaitTimeoutSeconds As Integer = 60
    Private Sub btnAddSupplier_Click(sender As Object, e As EventArgs) Handles btnAddSupplier.Click
        Dim supplierForm As New frmSuppliers()
        supplierForm.ShowDialog()
    End Sub

    Private Sub btnMaterialsDataEntry_Click(sender As Object, e As EventArgs) Handles btnMaterialsDataEntry.Click
        Dim fabricEntryForm As New frmFabricEntryForm()
        fabricEntryForm.ShowDialog() ' Use Show() if you want it non-modal
    End Sub

    Private Sub btnAddModels_Click(sender As Object, e As EventArgs) Handles btnAddModels.Click
        Dim ModelInformationForm As New frmAddModelInformation()
        ModelInformationForm.ShowDialog()
    End Sub



    Private Sub frmDashboard_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    ' ------------------------------------------------------------
    ' Sub: btnRegenerateSchema_Click
    ' What it does:
    '   When the user clicks the button, it:
    '     1) Verifies the Task Scheduler task exists
    '     2) Invokes the task immediately
    '     3) Optionally displays last run time after invocation
    ' Parameters:
    '   sender, e : standard WinForms event args
    ' Depends on:
    '   - ScheduledTaskHelper.TaskExists
    '   - ScheduledTaskHelper.RunScheduledTask
    '   - ScheduledTaskHelper.GetTaskLastRunTime (optional)
    ' ------------------------------------------------------------
    Private Sub btnRefreshSchema_Click(sender As Object, e As EventArgs) Handles btnRefreshSchema.Click
        Try
            Dim previousUtc As DateTime? = Nothing
            If File.Exists(SchemaJsonPath) Then
                previousUtc = File.GetLastWriteTimeUtc(SchemaJsonPath)
            End If

            ' 1) Confirm the task exists
            If Not TaskExists(SchemaTaskName) Then
                MessageBox.Show($"Scheduled task not found: {SchemaTaskName}" & Environment.NewLine &
                                "Open Task Scheduler and verify the exact task name (and folder path, if any).",
                                "Task Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If

            ' 2) Run the task
            If Not RunScheduledTask(SchemaTaskName) Then
                MessageBox.Show("Failed to start the scheduled task." & Environment.NewLine &
                                "If it requires elevation, run this app as Administrator or set the task to 'Run with highest privileges'.",
                                "Start Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            ' 3) Wait for schema JSON file to be updated
            Dim updated As Boolean = WaitForFileChange(SchemaJsonPath, previousUtc, WaitTimeoutSeconds)

            ' 4) Show result
            If updated Then
                Dim newTime As DateTime = File.GetLastWriteTime(SchemaJsonPath)
                Dim msg As String = $"Schema updated: {SchemaJsonPath}{Environment.NewLine}Last Write: {newTime}"
                MessageBox.Show(msg, "Schema Refreshed", MessageBoxButtons.OK, MessageBoxIcon.Information)

                ' Optionally reload a preview TextBox if it exists
                ReloadSchemaPreviewIfPresent()
            Else
                MessageBox.Show($"Timed out waiting for schema update after {WaitTimeoutSeconds} seconds." & Environment.NewLine &
                                $"Expected file: {SchemaJsonPath}",
                                "Timeout", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

        Catch ex As Exception
            MessageBox.Show("Unexpected error while refreshing schema:" & Environment.NewLine & ex.Message,
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ------------------------------------------------------------
    ' Sub: ReloadSchemaPreviewIfPresent
    ' Purpose:
    '   If a TextBox named txtSchemaPreview exists, load the JSON file.
    '   If a Label named lblSchemaStatus exists, update it with status.
    ' ------------------------------------------------------------
    Private Sub ReloadSchemaPreviewIfPresent()
        Try
            ' TextBox preview
            Dim ctl = Me.Controls.Find("txtSchemaPreview", True)
            If ctl IsNot Nothing AndAlso ctl.Length > 0 AndAlso TypeOf ctl(0) Is TextBox Then
                Dim tb As TextBox = DirectCast(ctl(0), TextBox)
                If File.Exists(SchemaJsonPath) Then
                    tb.Text = File.ReadAllText(SchemaJsonPath)
                Else
                    tb.Text = "(Schema file not found)"
                End If
            End If

            ' Status label
            Dim lbl = Me.Controls.Find("lblSchemaStatus", True)
            If lbl IsNot Nothing AndAlso lbl.Length > 0 AndAlso TypeOf lbl(0) Is Label Then
                Dim L As Label = DirectCast(lbl(0), Label)
                If File.Exists(SchemaJsonPath) Then
                    L.Text = "Schema refreshed: " & File.GetLastWriteTime(SchemaJsonPath).ToString()
                Else
                    L.Text = "Schema file not found."
                End If
            End If

        Catch
            ' Ignore preview errors
        End Try
    End Sub

End Class