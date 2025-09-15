Imports System

Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.Json
Imports System.Windows.Forms
Imports System.IO
Imports System.IO.Compression


Public Class frmDashboard

    ' ---- Update these constants as needed ----
    Private Const SchemaTaskName As String = "export_vbGGC_DailySchemaToJson"
    Private ReadOnly SchemaJsonPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Scripts\Schema\DailySchema.json")
    Private Const WaitTimeoutSeconds As Integer = 60

    Private RepoRoot As String = "C:\Users\ksmar\visualStudioProjects\GGC"  ' e.g., C:\Users\you\source\repos\YourRepo
    Private RepoOwner As String = "ksmartz"     ' e.g., dustcoversforyou
    Private RepoName As String = "GigGearCovers"               ' e.g., GGC
    Private RepoBranch As String = "master"                       ' e.g., main or master
    Private IncludeExtensions As String() = {
        ".vb", ".sln", ".vbproj", ".resx", ".config", ".json", ".ps1", ".csproj"
    }

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



    ' =================================================================================================
    ' Sub: frmDashboard_Load
    ' Date: 2025-09-06
    ' Purpose:
    '   Initialize the ListView columns (once) and configure view style.
    ' Dependencies:
    '   Requires lvGitLinks to exist on the form.
    ' Change Log:
    '   2025-09-06: Replaced "Is 0" with "= 0" in column-check logic.
    ' =================================================================================================
    Private Sub frmDashboard_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Ensure ListView uses Details view for columns
        lvGitLinks.View = View.Details
        lvGitLinks.FullRowSelect = True
        lvGitLinks.HideSelection = False

        ' ✅ Correct: compare Integer count with "=" instead of "Is"
        If lvGitLinks.Columns.Count = 0 Then
            lvGitLinks.Columns.Add("File", 420)     ' File display text (relative path)
            lvGitLinks.Columns.Add("URL", 600)      ' Full GitHub URL
        End If
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

    ' ================================================================
    ' File: frmDashboard.vb
    ' Purpose:
    '   Provide a button (btnGenerateGitLinks) that scans the local repo
    '   and generates clickable GitHub links for files (classes, modules,
    '   forms, projects, etc.). Links are shown in lvGitLinks and can be:
    '     - Opened via double-click
    '     - Copied via btnCopyGitLinks
    ' Dependencies:
    '   - HelperFunctions.BuildGitHubLinks / BuildGitHubBlobUrl
    '   - Controls expected on the form:
    '       * Button  : btnGenerateGitLinks
    '       * ListView: lvGitLinks (View=Details; Columns: "File", "GitHub URL")
    '       * Button  : btnCopyGitLinks (optional convenience)
    ' Notes:
    '   - Set the constants RepoRoot/Owner/Repo/Branch for your environment
    ' ================================================================



    ' ------------------------------------------------------------
    ' Overridden OnLoad:
    '   Wire up events with AddHandler so we don't rely on WithEvents.
    '   This avoids BC30506 when the Designer hasn't created those fields.
    ' ------------------------------------------------------------

    'btnGenerateGitLinks
    Protected Overrides Sub OnLoad(e As EventArgs)
        MyBase.OnLoad(e)

        ' Defensive: verify the controls exist on the form
        If btnCopyGitLinks IsNot Nothing Then
            RemoveHandler btnCopyGitLinks.Click, AddressOf btnCopyGitLinks_Click
            AddHandler btnCopyGitLinks.Click, AddressOf btnCopyGitLinks_Click
        End If

        ' Do NOT wire btnGenerateGitLinks_Click or lvGitLinks_DoubleClick here
        ' because those procedures already use "Handles ..."
    End Sub
    ' ------------------------------------------------------------
    ' Sub: EnsureUiControls
    ' Purpose:
    '   Find existing controls by name; if missing, create them.
    ' ------------------------------------------------------------
    Private Sub EnsureUiControls()
        ' Try to find existing (designer) controls
        Dim foundGen = Me.Controls.Find("btnGenerateGitLinks", True)
        If foundGen IsNot Nothing AndAlso foundGen.Length > 0 AndAlso TypeOf foundGen(0) Is Button Then
            btnGenerateGitLinks = DirectCast(foundGen(0), Button)
        End If

        Dim foundCopy = Me.Controls.Find("btnCopyGitLinks", True)
        If foundCopy IsNot Nothing AndAlso foundCopy.Length > 0 AndAlso TypeOf foundCopy(0) Is Button Then
            btnCopyGitLinks = DirectCast(foundCopy(0), Button)
        End If

        Dim foundLv = Me.Controls.Find("lvGitLinks", True)
        ' ✅ Correct (Length uses > 0, no Is against Integer)
        If foundLv.Length > 0 AndAlso TypeOf foundLv(0) Is ListView Then
            lvGitLinks = DirectCast(foundLv(0), ListView)
        End If


        ' Create them if missing
        If btnGenerateGitLinks Is Nothing Then
            btnGenerateGitLinks = New Button() With {
                .Name = "btnGenerateGitLinks",
                .Text = "Generate Git Links",
                .AutoSize = True,
                .Left = 12,
                .Top = 12
            }
            Me.Controls.Add(btnGenerateGitLinks)
        End If

        If btnCopyGitLinks Is Nothing Then
            btnCopyGitLinks = New Button() With {
                .Name = "btnCopyGitLinks",
                .Text = "Copy All Links",
                .AutoSize = True,
                .Left = btnGenerateGitLinks.Left + btnGenerateGitLinks.Width + 12,
                .Top = 12
            }
            Me.Controls.Add(btnCopyGitLinks)
        End If

        If lvGitLinks Is Nothing Then
            lvGitLinks = New ListView() With {
                .Name = "lvGitLinks",
                .Left = 12,
                .Top = btnGenerateGitLinks.Bottom + 12,
                .Width = Me.ClientSize.Width - 24,
                .Height = Me.ClientSize.Height - (btnGenerateGitLinks.Bottom + 24),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom,
                .View = View.Details,
                .FullRowSelect = True
            }
            Me.Controls.Add(lvGitLinks)
        End If
    End Sub

    ' ------------------------------------------------------------
    ' Sub: WireEvents
    ' Purpose:
    '   Attach handlers with AddHandler (avoids Handles/WithEvents issues).
    ' ------------------------------------------------------------
    Private Sub WireEvents()
        RemoveHandler btnGenerateGitLinks.Click, AddressOf btnGenerateGitLinks_Click
        AddHandler btnGenerateGitLinks.Click, AddressOf btnGenerateGitLinks_Click

        RemoveHandler btnCopyGitLinks.Click, AddressOf btnCopyGitLinks_Click
        AddHandler btnCopyGitLinks.Click, AddressOf btnCopyGitLinks_Click

        RemoveHandler lvGitLinks.DoubleClick, AddressOf lvGitLinks_DoubleClick
        AddHandler lvGitLinks.DoubleClick, AddressOf lvGitLinks_DoubleClick
    End Sub

    ' =================================================================================================
    ' Sub: btnGenerateGitLinks_Click
    ' Date: 2025-09-06 (revised)
    ' Purpose:
    '   Validate/select the local repo folder, build GitHub links, and populate lvGitLinks.
    ' Notes:
    '   Adds diagnostics so you immediately see whether files were found.
    ' Dependencies:
    '   - GitHubLinkHelpers.BuildGitHubLinks
    '   - ValidateAndSelectRepoRoot (below)
    ' =================================================================================================
    Private Sub btnGenerateGitLinks_Click(sender As Object, e As EventArgs) Handles btnGenerateGitLinks.Click
        Try
            ' Ensure we have a real folder to scan
            If Not ValidateAndSelectRepoRoot() Then
                MessageBox.Show("No valid repository folder selected. Cannot generate links.", "Repo Required",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If lvGitLinks Is Nothing Then
                MessageBox.Show("ListView 'lvGitLinks' not found on the form.", "Missing Control",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            lvGitLinks.BeginUpdate()
            lvGitLinks.Items.Clear()

            Dim links As List(Of GitHubLinkHelpers.GitLink) =
            GitHubLinkHelpers.BuildGitHubLinks(RepoRoot, RepoOwner, RepoName, RepoBranch, IncludeExtensions)

            ' Quick diagnostics to help you debug
            Dim count As Integer = If(links Is Nothing, 0, links.Count)
            If count = 0 Then
                MessageBox.Show(
                "No matching files were found." & Environment.NewLine &
                "• RepoRoot: " & RepoRoot & Environment.NewLine &
                "• RepoOwner/RepoName/Branch: " & RepoOwner & "/" & RepoName & " (" & RepoBranch & ")" & Environment.NewLine &
                "• IncludeExtensions: " & String.Join(", ", IncludeExtensions),
                "0 Results", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

            For Each link As GitHubLinkHelpers.GitLink In links
                Dim item As New ListViewItem(link.Text) ' Relative path
                item.SubItems.Add(link.Url)             ' GitHub URL
                lvGitLinks.Items.Add(item)
            Next

        Catch ex As Exception
            MessageBox.Show($"Error generating links: {ex.Message}", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If lvGitLinks IsNot Nothing Then lvGitLinks.EndUpdate()
        End Try
    End Sub

    ' =================================================================================================
    ' Function: ValidateAndSelectRepoRoot
    ' Date: 2025-09-06
    ' Purpose:
    '   Ensure RepoRoot points to an existing local folder. If not, prompt the user once to pick it.
    ' Returns:
    '   True if RepoRoot exists after validation/selection; otherwise False.
    ' Notes:
    '   This is a one-time helper so you don't silently get 0 results when RepoRoot is wrong.
    ' =================================================================================================
    Private Function ValidateAndSelectRepoRoot() As Boolean
        Try
            ' If the configured path exists, we're good
            If Not String.IsNullOrWhiteSpace(RepoRoot) AndAlso Directory.Exists(RepoRoot) Then
                Return True
            End If

            ' Suggest a likely default based on your repo name if missing
            Dim guess As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                                           "Visual Studio Projects", "GGC")

            Using fbd As New FolderBrowserDialog()
                fbd.Description = "Select the LOCAL folder for your Git repo (the folder that contains GGC.sln)."
                fbd.ShowNewFolderButton = False
                fbd.SelectedPath = If(Directory.Exists(RepoRoot), RepoRoot,
                               If(Directory.Exists(guess), guess, Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)))

                If fbd.ShowDialog(Me) = DialogResult.OK AndAlso Directory.Exists(fbd.SelectedPath) Then
                    RepoRoot = fbd.SelectedPath
                    Return True
                End If
            End Using

        Catch
            ' Ignore and fall through
        End Try

        Return False
    End Function


    ' =================================================================================================
    ' Sub: lvGitLinks_DoubleClick
    ' Date: 2025-09-06
    ' Purpose:
    '   Open the GitHub URL for the selected item in the default browser.
    ' Dependencies:
    '   Requires that the second subitem contains the URL.
    ' Change Log:
    '   2025-09-06: Replaced "Is 0" with "= 0" in the selection count check.
    ' =================================================================================================
    Private Sub lvGitLinks_DoubleClick(sender As Object, e As EventArgs) Handles lvGitLinks.DoubleClick
        If lvGitLinks.SelectedItems.Count = 0 Then Return

        Dim sel As ListViewItem = lvGitLinks.SelectedItems(0)
        If sel.SubItems.Count < 2 Then Return

        Dim url As String = sel.SubItems(1).Text
        If String.IsNullOrWhiteSpace(url) Then Return

        Try
            Dim psi As New ProcessStartInfo(url) With {.UseShellExecute = True}
            Process.Start(psi)
        Catch ex As Exception
            MessageBox.Show($"Unable to open link: {ex.Message}", "Open URL", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub
    ' ------------------------------------------------------------
    ' Sub: btnCopyGitLinks_Click
    ' Purpose:
    '   Copy all generated GitHub URLs (from lvGitLinks) to the clipboard.
    ' ------------------------------------------------------------

    Private Sub btnCopyGitLinks_Click(sender As Object, e As EventArgs)
        Try
            If lvGitLinks Is Nothing OrElse lvGitLinks.Items.Count = 0 Then
                MessageBox.Show("No links to copy.", "Copied", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim sb As New StringBuilder()
            For Each it As ListViewItem In lvGitLinks.Items
                If it.SubItems.Count > 1 Then
                    sb.AppendLine(it.SubItems(1).Text)
                End If
            Next

            If sb.Length > 0 Then
                Clipboard.SetText(sb.ToString())
                MessageBox.Show("Copied all GitHub links to clipboard.", "Copied", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("No links to copy.", "Copied", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("Copy failed:" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Const SW_SHOWMINNOACTIVE As Integer = 7

    <DllImport("shell32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Shared Function ShellExecute(hwnd As IntPtr,
                                         lpOperation As String,
                                         lpFile As String,
                                         lpParameters As String,
                                         lpDirectory As String,
                                         nShowCmd As Integer) As IntPtr
    End Function

    ''' <summary>
    ''' Opens a folder in Explorer minimized (background) to avoid stealing focus.
    ''' </summary>
    Private Sub OpenFolderInBackground(folderPath As String)
        Try
            If String.IsNullOrWhiteSpace(folderPath) OrElse Not Directory.Exists(folderPath) Then Exit Sub
            ' Use ShellExecute with SW_SHOWMINNOACTIVE to open minimized without activating the window
            ShellExecute(IntPtr.Zero, "open", folderPath, Nothing, Nothing, SW_SHOWMINNOACTIVE)
        Catch
            ' Non-fatal if Explorer fails to open; ignore.
        End Try
    End Sub
    ' ================================================================================================
    ' Sub: btnExportClasses_Click
    ' Where: frmDashboard
    ' Purpose: Export VB classes (or full repo) to a ZIP, with strong validation but no ZipArchive usage.
    ' Date: 2025-09-10
    ' ================================================================================================
    Private Sub btnExportClasses_Click(sender As Object, e As EventArgs) Handles btnExportClasses.Click
        Try
            ' 1) Ask for repo root (folder that contains GGC.sln)
            Dim repoRoot As String = Nothing
            Using fbd As New FolderBrowserDialog()
                fbd.Description = "Select your GGC repository root (folder that contains GGC.sln)."
                fbd.ShowNewFolderButton = False
                fbd.SelectedPath = AppDomain.CurrentDomain.BaseDirectory
                If fbd.ShowDialog(Me) <> DialogResult.OK Then Exit Sub
                repoRoot = fbd.SelectedPath
            End Using

            If String.IsNullOrWhiteSpace(repoRoot) OrElse Not Directory.Exists(repoRoot) Then
                MessageBox.Show(Me, "Invalid folder selected.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If

            ' 2) Mode prompt: Yes = Classes-only ZIP (recommended); No = Full-repo ZIP
            Dim modeMsg As String =
            "Create a ZIP of ONLY your VB classes as .txt (recommended)?" & Environment.NewLine &
            "Choose 'No' to zip the entire repo (filtered)."
            Dim mode As CodeExporter.ExportMode =
            If(MessageBox.Show(Me, modeMsg, "Choose Export Mode",
                               MessageBoxButtons.YesNo,
                               MessageBoxIcon.Question) = DialogResult.Yes,
               CodeExporter.ExportMode.ClassesOnlyToTxtAndZip,
               CodeExporter.ExportMode.FullRepoZip)

            ' 3) Pre-check: make sure the chosen mode will actually have content
            If mode = CodeExporter.ExportMode.ClassesOnlyToTxtAndZip Then
                If Not RepoContainsVbFiles(repoRoot) Then
                    MessageBox.Show(Me,
                                "No .vb files were found beneath the folder you selected." & Environment.NewLine &
                                "Pick the folder that contains your GGC.sln (the solution's root).",
                                "Nothing to export",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)
                    Exit Sub
                End If
            Else
                If Not RepoContainsAnyFilesExcludingBuildDirs(repoRoot) Then
                    MessageBox.Show(Me,
                                "No exportable files were found under the selected folder (after excluding bin/obj/.git/.vs/etc.).",
                                "Nothing to export",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)
                    Exit Sub
                End If
            End If

            ' 4) Run export + schema refresh + zip
            Dim zipPath As String = Nothing
            Dim ok As Boolean = CodeExporter.RunExportWithSchemaZip(repoRoot, mode, zipPath)

            If Not ok OrElse String.IsNullOrWhiteSpace(zipPath) OrElse Not File.Exists(zipPath) Then
                MessageBox.Show(Me, "Exporter did not produce a zip file.", "Export failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            ' 5) Light sanity: the minimum valid ZIP (EOCD only) is ~22 bytes. Use a higher floor to guard empties.
            If Not ArchiveHasEntries(zipPath) Then
                MessageBox.Show(Me,
                            "The zip was created but appears to contain no files." & Environment.NewLine &
                            "Re-check the selected folder. You may have picked the wrong root.",
                            "Empty zip",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning)
                Exit Sub
            End If

            ' 6) Open the folder (background) and show a passive toast
            Dim zipFolder As String = Path.GetDirectoryName(zipPath)
            OpenFolderInBackground(zipFolder)
            NotifyExportCompleted(zipPath)

        Catch ex As Exception
            MessageBox.Show(Me,
                        "Export failed: " & ex.Message,
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error)
        End Try
    End Sub


    ' ================================================================================================
    ' Function: ArchiveHasEntries
    ' Purpose: Heuristic check that avoids System.IO.Compression types.
    '   - Returns False if the zip doesn't exist or is suspiciously tiny.
    '   - Paired with pre-checks so we don't rely on ZIP parsing at all.
    ' Date: 2025-09-10
    ' ================================================================================================
    Private Function ArchiveHasEntries(zipPath As String) As Boolean
        Try
            Dim fi As New FileInfo(zipPath)
            If Not fi.Exists Then Return False
            ' A totally empty ZIP (or bad write) will be tiny. Use a conservative floor.
            ' Most real exports will be >> 1 KB.
            Return fi.Length >= 200
        Catch
            Return False
        End Try
    End Function

    ' ================================================================================================
    ' Function: RepoContainsAnyFilesExcludingBuildDirs
    ' Purpose: Detect if there are any files worth zipping, skipping heavy/build/VCS/editor folders.
    ' Date: 2025-09-10
    ' ================================================================================================
    Private Function RepoContainsAnyFilesExcludingBuildDirs(root As String) As Boolean
        If String.IsNullOrWhiteSpace(root) OrElse Not Directory.Exists(root) Then Return False

        Dim excludedParts As String() = {"\bin\", "\obj\", "\.git\", "\.vs\", "\.idea\", "\node_modules\"}
        For Each file In Directory.EnumerateFiles(root, "*.*", SearchOption.AllDirectories)
            Dim p = file.Replace("/"c, "\"c)
            Dim skip As Boolean = excludedParts.Any(Function(part) p.IndexOf(part, StringComparison.OrdinalIgnoreCase) >= 0)
            If Not skip Then Return True
        Next
        Return False
    End Function


    ' ================================================================================================
    ' Function: RepoContainsVbFiles
    ' Purpose: Returns True if any *.vb files exist beneath the selected root.
    ' Date: 2025-09-10
    ' ================================================================================================
    Private Function RepoContainsVbFiles(root As String) As Boolean
        If String.IsNullOrWhiteSpace(root) OrElse Not Directory.Exists(root) Then Return False
        Return Directory.EnumerateFiles(root, "*.vb", SearchOption.AllDirectories).Any()
    End Function

    ''' <summary>
    ''' Non-blocking confirmation that shows the zip file name; does not ask anything.
    ''' </summary>
    Private Sub NotifyExportCompleted(zipPath As String)
        Try
            Dim fileName = Path.GetFileName(zipPath)
            ' Short, informational message that won't block your flow.
            ' If you prefer zero UI, remove this entirely.
            Dim msg = $"Export completed: {fileName}" & Environment.NewLine &
                      $"Location: {Path.GetDirectoryName(zipPath)}"
            MessageBox.Show(Me, msg, "Export Ready", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch
            ' Ignore UI errors
        End Try
    End Sub

End Class





