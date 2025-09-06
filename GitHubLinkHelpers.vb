
' =================================================================================================
' Module: GitHubLinkHelpers
' Date: 2025-09-06
' Purpose:
'   Provide strongly-typed helpers for building GitHub blob URLs from a local repo clone.
' Notes:
'   An optional overload accepts an "extraExcludes" list so callers can pass additional
'   folders/files to skip without relying on a global variable (prevents BC30451).
' =================================================================================================
Imports System.IO

Public Module GitHubLinkHelpers
    ' ---------------------------------------------------------------------------------------------
    ' DTO for a link row (safer than tuples for broad VB.NET compatibility)
    ' ---------------------------------------------------------------------------------------------
    Public Class GitLink
        Public Property Text As String  ' Relative path shown to the user
        Public Property Url As String   ' GitHub blob URL
        Public Sub New(text As String, url As String)
            Me.Text = text
            Me.Url = url
        End Sub
    End Class

    ' =================================================================================================
    ' Function: BuildGitHubLinks
    ' Date: 2025-09-06
    ' Purpose:
    '   Enumerate files from RepoRoot, filter by IncludeExtensions, skip common excludes,
    '   and return GitHub blob URLs.
    ' Dependencies:
    '   System.IO, System.Linq
    ' =================================================================================================
    Public Function BuildGitHubLinks(repoRoot As String,
                                     owner As String,
                                     repo As String,
                                     branch As String,
                                     includeExtensions As IEnumerable(Of String)) As List(Of GitLink)

        ' Call the overload with an empty "extraExcludes" to keep the logic in one place
        Return BuildGitHubLinks(repoRoot, owner, repo, branch, includeExtensions, Enumerable.Empty(Of String)())
    End Function

    ' =================================================================================================
    ' Function: BuildGitHubLinks (overload with excludes)
    ' Date: 2025-09-06
    ' Purpose:
    '   Same as above, but allows the caller to pass extra excludes. This avoids relying on any
    '   undeclared "ExtraExcludes" variable (fix for BC30451).
    ' =================================================================================================
    Public Function BuildGitHubLinks(repoRoot As String,
                                     owner As String,
                                     repo As String,
                                     branch As String,
                                     includeExtensions As IEnumerable(Of String),
                                     extraExcludes As IEnumerable(Of String)) As List(Of GitLink)

        Dim results As New List(Of GitLink)

        If String.IsNullOrWhiteSpace(repoRoot) OrElse Not Directory.Exists(repoRoot) Then
            Return results ' Nothing to do if root is missing
        End If

        ' Normalize filters
        Dim extSet As HashSet(Of String) =
            New HashSet(Of String)(includeExtensions.Select(Function(s) s.Trim().ToLowerInvariant()))

        ' Default excludes (bin/obj/.git) + caller-provided
        Dim defaultExcludes As String() = {".git", "bin", "obj", ".vs", ".idea", ".gitignore"}
        Dim allExcludes As HashSet(Of String) =
            New HashSet(Of String)(defaultExcludes.Concat(extraExcludes).Select(Function(s) s.Trim().ToLowerInvariant()))

        ' Enumerate files recursively
        For Each filePath In Directory.EnumerateFiles(repoRoot, "*.*", SearchOption.AllDirectories)
            Dim rel As String = GetRelativePath(repoRoot, filePath)

            ' Skip excluded folders/files if the relative path contains an excluded segment
            Dim relLower = rel.Replace("\", "/").ToLowerInvariant()
            Dim skip As Boolean = allExcludes.Any(Function(ex) relLower.Contains("/" & ex & "/") OrElse relLower.EndsWith("/" & ex) OrElse relLower.Contains(ex & "/"))
            If skip Then Continue For

            ' Filter by extension
            Dim ext As String = Path.GetExtension(filePath).ToLowerInvariant()
            If extSet.Count > 0 AndAlso Not extSet.Contains(ext) Then Continue For

            ' Construct GitHub blob URL
            Dim url As String = $"https://github.com/{owner}/{repo}/blob/{branch}/{rel.Replace("\", "/")}"

            results.Add(New GitLink(rel, url))
        Next

        ' Sort by path for nicer presentation
        results.Sort(Function(a, b) String.Compare(a.Text, b.Text, StringComparison.OrdinalIgnoreCase))
        Return results
    End Function

    ' Small helper to compute a relative path in .NET Framework/.NET versions without Path.GetRelativePath
    Private Function GetRelativePath(root As String, fullPath As String) As String
        Dim rootFull = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        Dim fileFull = Path.GetFullPath(fullPath)
        If fileFull.StartsWith(rootFull, StringComparison.OrdinalIgnoreCase) Then
            Dim rel = fileFull.Substring(rootFull.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            Return rel
        End If
        Return Path.GetFileName(fullPath)
    End Function


End Module
