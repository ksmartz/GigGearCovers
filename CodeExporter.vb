' =================================================================================================
' File: CodeExporter.vb
' Purpose: Export VB.NET source files (*.vb) to .txt, refresh schema (PowerShell), copy DailySchema,
'          and package everything into a single ZIP file for upload/sharing.
' Dependencies:
'   - System.IO
'   - System.Text
'   - System.Data (only if you later extend for DB ops)
'   - System.IO.Compression (ZipFile)
'   - System.IO.Compression.FileSystem (for ZipFile.CreateFromDirectory)
'   - PowerShell on host to run exportschematojson.ps1
' Date: 2025-09-08
' -------------------------------------------------------------------------------------------------
' How it works (classes-only flow):
'   1) Optionally refresh the JSON schema by invoking Scripts\Schema\exportschematojson.ps1
'   2) Export all *.vb (excluding Designer/auto-gen, bin/obj/.git/.vs) to .txt under a timestamped folder
'   3) Copy the freshly generated DailySchema.json into the export folder (preserving relative subpath)
'   4) Zip that timestamped folder into a single .zip
'
' How it works (full repo flow):
'   1) Optionally refresh schema via the PowerShell script
'   2) Zip the entire repo folder (excluding bin/obj/.git/.vs by default to keep size sane)
'
' Caller chooses flow via a parameter.
' =================================================================================================
Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Diagnostics

Public Module CodeExporter

    Public Enum ExportMode
        ClassesOnlyToTxtAndZip = 0
        FullRepoZip = 1
    End Enum


    ' ------------------------------------------------------------------------------------------------
    ' Export VB -> TXT (preserving relative structure; excludes Designer/auto-generated + build/vcs dirs)
    ' ------------------------------------------------------------------------------------------------
    Public Sub ExportVbClassesToText(ByVal sourceRoot As String,
                                     ByVal exportRoot As String,
                                     ByRef filesExported As Integer)
        filesExported = 0
        If String.IsNullOrWhiteSpace(sourceRoot) OrElse Not Directory.Exists(sourceRoot) Then
            Throw New DirectoryNotFoundException("Source root not found for VB export.")
        End If
        Directory.CreateDirectory(exportRoot)

        Dim allVb = Directory.EnumerateFiles(sourceRoot, "*.vb", SearchOption.AllDirectories) _
                             .Where(Function(p) Not IsIgnoredPathForClassesOnly(p))

        Dim srcFull = Path.GetFullPath(sourceRoot).TrimEnd(Path.DirectorySeparatorChar)
        For Each vbPath In allVb
            Dim rel As String = GetRelativePath(srcFull, vbPath)
            Dim relTxt As String = Path.ChangeExtension(rel, ".txt")
            Dim outPath As String = Path.Combine(exportRoot, relTxt)
            Directory.CreateDirectory(Path.GetDirectoryName(outPath))

            Dim code As String = File.ReadAllText(vbPath, DetectEncodingOrDefault(vbPath))
            Dim header As String =
                $"' Exported from: {rel}{Environment.NewLine}" &
                $"' Exported at: {DateTime.Now}{Environment.NewLine}" &
                "' --------------------------------------------------------------------------------" & Environment.NewLine

            File.WriteAllText(outPath, header & code, New UTF8Encoding(encoderShouldEmitUTF8Identifier:=False))
            filesExported += 1
        Next
    End Sub

    ' ------------------------------------------------------------------------------------------------
    ' Schema refresh + copy helpers
    ' ------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Runs Scripts\Schema\exportschematojson.ps1 from the repo to refresh DailySchema.json.
    ''' Throws if the script reports a non-zero exit code.
    ''' </summary>
    Public Sub RefreshDailySchema(repoRoot As String)
        Dim scriptPath As String = Path.Combine(repoRoot, "Scripts", "Schema", "exportschematojson.ps1")
        If Not File.Exists(scriptPath) Then
            ' Don’t fail the entire export if the script is missing; just return quietly.
            Return
        End If

        ' Use PowerShell to execute the script
        Dim psi As New ProcessStartInfo() With {
            .FileName = "powershell.exe",
            .Arguments = "-NoProfile -ExecutionPolicy Bypass -File " & Quote(scriptPath),
            .UseShellExecute = False,
            .RedirectStandardOutput = True,
            .RedirectStandardError = True,
            .CreateNoWindow = True,
            .WorkingDirectory = Path.GetDirectoryName(scriptPath)
        }

        Using proc As Process = Process.Start(psi)
            Dim stdOut As String = proc.StandardOutput.ReadToEnd()
            Dim stdErr As String = proc.StandardError.ReadToEnd()
            proc.WaitForExit()

            If proc.ExitCode <> 0 Then
                Throw New InvalidOperationException(
                    $"Schema export script failed with exit code {proc.ExitCode}.{Environment.NewLine}STDOUT:{Environment.NewLine}{stdOut}{Environment.NewLine}STDERR:{Environment.NewLine}{stdErr}")
            End If
        End Using
    End Sub

    ''' <summary>
    ''' Copies DailySchema.json (if present) to the exportRoot, preserving Scripts\Schema subfolder.
    ''' </summary>
    Private Sub CopyRefreshedSchemaIfFound(repoRoot As String, exportRoot As String)
        Dim src As String = Path.Combine(repoRoot, "Scripts", "Schema", "DailySchema.json")
        If Not System.IO.File.Exists(src) Then Exit Sub

        Dim destDir As String = Path.Combine(exportRoot, "Scripts", "Schema")
        System.IO.Directory.CreateDirectory(destDir)

        Dim dest As String = Path.Combine(destDir, "DailySchema.json")
        ' Fully qualified to avoid binding to Microsoft.VisualBasic.Strings.Copy
        System.IO.File.Copy(src, dest, True)
    End Sub

    ' ---------------------------------------------------------------------------------------------------
    ' Sub: CopyDirectoryFiltered
    ' Module: CodeExporter
    ' Purpose: Recursively copy repo contents into a temp folder while skipping excluded paths/files.
    ' Dependencies: System.IO
    ' Date: 2025-09-08
    ' ---------------------------------------------------------------------------------------------------
    Private Sub CopyDirectoryFiltered(sourceDir As String,
                                  destDir As String,
                                  exclude As Func(Of String, Boolean))
        Dim stack As New Stack(Of String)
        stack.Push(sourceDir)

        Dim srcRoot As String = System.IO.Path.GetFullPath(sourceDir).TrimEnd(System.IO.Path.DirectorySeparatorChar)

        While stack.Count > 0
            Dim current As String = stack.Pop()
            Dim rel As String = GetRelativePath(srcRoot, current)
            Dim target As String = If(String.IsNullOrEmpty(rel), destDir, System.IO.Path.Combine(destDir, rel))

            ' Ensure the destination directory exists for the current folder
            System.IO.Directory.CreateDirectory(target)

            ' Copy files in the current directory (except excluded)
            For Each file In System.IO.Directory.GetFiles(current)
                If Not exclude(file) Then
                    Dim relFile As String = GetRelativePath(srcRoot, file)
                    Dim destFile As String = System.IO.Path.Combine(destDir, relFile)

                    ' Ensure the destination subdirectory exists for this file
                    Dim destDirPath As String = System.IO.Path.GetDirectoryName(destFile)
                    System.IO.Directory.CreateDirectory(destDirPath)

                    ' Copy file (overwrite := True)
                    System.IO.File.Copy(file, destFile, True)
                End If
            Next

            ' Recurse into subdirectories (except excluded)
            For Each subDir In System.IO.Directory.GetDirectories(current)
                If Not exclude(subDir) Then
                    stack.Push(subDir)
                End If
            Next
        End While
    End Sub



    Private Function IsExcludedForFullRepoZip(path As String) As Boolean
        Dim p = path.Replace("/"c, "\"c)

        ' Exclude heavy/build/scm/editor dirs
        If p.IndexOf("\bin\", StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        If p.IndexOf("\obj\", StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        If p.IndexOf("\.git\", StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        If p.IndexOf("\.vs\", StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        If p.IndexOf("\.idea\", StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        If p.IndexOf("\node_modules\", StringComparison.OrdinalIgnoreCase) >= 0 Then Return True

        ' Optionally skip large media: uncomment if needed
        ' Dim ext = Path.GetExtension(p)
        ' If {".png", ".jpg", ".jpeg", ".gif", ".mp4", ".mov", ".zip"}.Contains(ext.ToLowerInvariant()) Then Return True

        Return False
    End Function

    ' ------------------------------------------------------------------------------------------------
    ' Shared utilities
    ' ------------------------------------------------------------------------------------------------
    ' ---------------------------------------------------------------------------------------------------
    ' Function: IsIgnoredPathForClassesOnly
    ' Module: CodeExporter
    ' Purpose: Decide whether a file should be excluded from the classes-only TXT export.
    ' Dependencies: System.IO
    ' Date: 2025-09-08
    ' ---------------------------------------------------------------------------------------------------
    Private Function IsIgnoredPathForClassesOnly(fullPath As String) As Boolean
        Dim p As String = fullPath.Replace("/"c, "\"c)

        ' Exclude build and VCS folders
        If p.IndexOf("\bin\", StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        If p.IndexOf("\obj\", StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        If p.IndexOf("\.git\", StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        If p.IndexOf("\.vs\", StringComparison.OrdinalIgnoreCase) >= 0 Then Return True

        ' Exclude common auto-generated VB files
        Dim fileName As String = System.IO.Path.GetFileName(p)
        If fileName.EndsWith(".Designer.vb", StringComparison.OrdinalIgnoreCase) Then Return True
        If fileName.Equals("Resources.Designer.vb", StringComparison.OrdinalIgnoreCase) Then Return True
        If fileName.Equals("Settings.Designer.vb", StringComparison.OrdinalIgnoreCase) Then Return True
        If fileName.Equals("Application.Designer.vb", StringComparison.OrdinalIgnoreCase) Then Return True
        If fileName.Equals("AssemblyInfo.vb", StringComparison.OrdinalIgnoreCase) Then Return True ' usually noise

        Return False
    End Function


    Private Function GetRelativePath(root As String, fullPath As String) As String
        Dim f = Path.GetFullPath(fullPath)
        If Not f.StartsWith(root, StringComparison.OrdinalIgnoreCase) Then
            Return Path.GetFileName(f)
        End If
        Dim rel = f.Substring(root.Length)
        If rel.StartsWith(Path.DirectorySeparatorChar) Then rel = rel.Substring(1)
        Return rel
    End Function

    Private Function DetectEncodingOrDefault(filePath As String) As Encoding
        Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            If fs.Length >= 3 Then
                Dim bom(2) As Byte
                fs.Read(bom, 0, 3)
                If bom(0) = &HEF AndAlso bom(1) = &HBB AndAlso bom(2) = &HBF Then
                    Return New UTF8Encoding(True) ' UTF-8 with BOM
                End If
                fs.Position = 0
            End If
        End Using
        Return New UTF8Encoding(False) ' UTF-8 no BOM
    End Function

    Private Function Quote(path As String) As String
        If String.IsNullOrEmpty(path) Then Return """"""""
        If path.Contains("""") Then path = path.Replace("""", "\""")
        Return """" & path & """"
    End Function

    Private Sub SafeDeleteDirectory(dir As String)
        Try
            If Directory.Exists(dir) Then Directory.Delete(dir, recursive:=True)
        Catch
            ' ignore temp cleanup errors
        End Try
    End Sub
    ' ---------------------------------------------------------------------------------------------------
    ' Function: RunExportWithSchemaZip
    ' Module: CodeExporter
    ' Purpose: Refresh schema; then either export classes-to-TXT and zip (+ DailySchema.json),
    '          or create a filtered full-repo ZIP. Returns True on success and outputs the ZIP path.
    ' Dependencies: System.IO, System.IO.Compression, System.Diagnostics
    ' Date: 2025-09-08 (replaced duplicate definitions to resolve BC30269)
    ' ---------------------------------------------------------------------------------------------------
    Public Function RunExportWithSchemaZip(repoRoot As String,
                                       mode As ExportMode,
                                       ByRef zipPathOut As String) As Boolean
        If String.IsNullOrWhiteSpace(repoRoot) OrElse Not System.IO.Directory.Exists(repoRoot) Then
            Throw New System.IO.DirectoryNotFoundException("Repository root not found.")
        End If

        ' 1) Refresh schema first (non-fatal if script missing; handled inside helper)
        RefreshDailySchema(repoRoot)

        ' 2) Prepare export parent folder under Documents
        Dim stamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
        Dim docs As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim exportParent As String = System.IO.Path.Combine(docs, "GGC_CodeExports")
        System.IO.Directory.CreateDirectory(exportParent)

        If mode = ExportMode.FullRepoZip Then
            ' --- Full repo ZIP ---
            Dim tempCopy As String = System.IO.Path.Combine(exportParent, $"RepoSnapshot_{stamp}")
            System.IO.Directory.CreateDirectory(tempCopy)

            ' Copy repo into temp (filtered)
            CopyDirectoryFiltered(repoRoot, tempCopy, AddressOf IsExcludedForFullRepoZip)

            ' Zip the temp snapshot
            Dim zipPath As String = System.IO.Path.Combine(exportParent, $"GGC_FullRepo_{stamp}.zip")
            If System.IO.File.Exists(zipPath) Then System.IO.File.Delete(zipPath)
            System.IO.Compression.ZipFile.CreateFromDirectory(tempCopy, zipPath,
                                                          System.IO.Compression.CompressionLevel.Optimal,
                                                          includeBaseDirectory:=False)

            ' Cleanup temp
            SafeDeleteDirectory(tempCopy)

            zipPathOut = zipPath

            ' If you log/display the name, use Path.GetFileName (static), not string.GetFileName
            Dim fileName As String = System.IO.Path.GetFileName(zipPathOut)

            Return True

        Else
            ' --- Classes-only to TXT then ZIP ---
            Dim exportRoot As String = System.IO.Path.Combine(exportParent, $"Export_{stamp}")

            ' Export *.vb -> *.txt (filtered)
            Dim count As Integer = 0
            ExportVbClassesToText(repoRoot, exportRoot, count)

            ' Copy refreshed schema file if present
            CopyRefreshedSchemaIfFound(repoRoot, exportRoot)

            ' Zip the export directory
            Dim zipPath As String = System.IO.Path.Combine(exportParent, $"GGC_Classes_{stamp}.zip")
            If System.IO.File.Exists(zipPath) Then System.IO.File.Delete(zipPath)
            System.IO.Compression.ZipFile.CreateFromDirectory(exportRoot, zipPath,
                                                          System.IO.Compression.CompressionLevel.Optimal,
                                                          includeBaseDirectory:=False)

            zipPathOut = zipPath

            ' If you log/display the name, use Path.GetFileName (static), not string.GetFileName
            Dim fileName As String = System.IO.Path.GetFileName(zipPathOut)

            Return True
        End If
    End Function

End Module
