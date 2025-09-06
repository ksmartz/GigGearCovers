Imports System
Imports System.Diagnostics
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Module HelperFunctions



    Public Function CalculateRetailPrice(grandTotal As Decimal, profit As Decimal) As Decimal
        Dim price As Decimal = grandTotal + profit
        ' Round up to the next .95
        Dim rounded As Decimal = Math.Ceiling(price - 0.95D) + 0.95D
        Return rounded
    End Function

    Public Function CalculateFabricWeightInOunces(weightPerLinearYard As Decimal, fabricRollWidth As Decimal, totalFabricSquareInches As Decimal) As Decimal
        If fabricRollWidth <= 0 Then Return 0D
        Dim weightPerSqInch As Decimal = weightPerLinearYard / (36D * fabricRollWidth)
        Return Math.Round(weightPerSqInch * totalFabricSquareInches, 2)
    End Function
    ' In HelperFunctions.vb or frmAddModelInformation.vb

    Public Function CalculateModelMaterialWeights(totalFabricSquareInches As Decimal) As (
    weight_PaddingOnly As Decimal?,
    weight_ChoiceWaterproof As Decimal?,
    weight_ChoiceWaterproof_Padded As Decimal?,
    weight_PremiumSyntheticLeather As Decimal?,
    weight_PremiumSyntheticLeather_Padded As Decimal?
)
        Dim db As New DbConnectionManager()

        Dim rowCW = db.GetActiveFabricBrandProductName("Choice Waterproof")
        Dim rowSL = db.GetActiveFabricBrandProductName("Premium Synthetic Leather")
        Dim rowPad = db.GetActiveFabricBrandProductName("Padding")

        Dim weightCW As Decimal? = Nothing
        Dim weightSL As Decimal? = Nothing
        Dim weightPad As Decimal? = Nothing
        Dim weightCWPad As Decimal? = Nothing
        Dim weightSLPad As Decimal? = Nothing

        If rowCW IsNot Nothing Then
            weightCW = CalculateFabricWeightInOunces(
            Convert.ToDecimal(rowCW("WeightPerLinearYard")),
            Convert.ToDecimal(rowCW("FabricRollWidth")),
            totalFabricSquareInches)
        End If
        If rowSL IsNot Nothing Then
            weightSL = CalculateFabricWeightInOunces(
            Convert.ToDecimal(rowSL("WeightPerLinearYard")),
            Convert.ToDecimal(rowSL("FabricRollWidth")),
            totalFabricSquareInches)
        End If
        If rowPad IsNot Nothing Then
            weightPad = CalculateFabricWeightInOunces(
            Convert.ToDecimal(rowPad("WeightPerLinearYard")),
            Convert.ToDecimal(rowPad("FabricRollWidth")),
            totalFabricSquareInches)
        End If

        If weightCW.HasValue AndAlso weightPad.HasValue Then
            weightCWPad = Math.Round(weightCW.Value + weightPad.Value, 2)
        End If
        If weightSL.HasValue AndAlso weightPad.HasValue Then
            weightSLPad = Math.Round(weightSL.Value + weightPad.Value, 2)
        End If

        Return (weightPad, weightCW, weightCWPad, weightSL, weightSLPad)
    End Function
    ' Place this in your form or a shared module as needed
    Public Sub CalculateAndInsertModelLaborCosts(
    modelId As Integer,
    totalFabricSquareInches As Decimal,
    costs As (costPerSqInch_ChoiceWaterproof As Decimal?, costPerSqInch_PremiumSyntheticLeather As Decimal?, costPerSqInch_Padding As Decimal?, baseCost_ChoiceWaterproof As Decimal?, baseCost_PremiumSyntheticLeather As Decimal?, baseCost_ChoiceWaterproof_Padded As Decimal?, baseCost_PremiumSyntheticLeather_Padded As Decimal?, baseCost_PaddingOnly As Decimal?),
    weights As (weight_PaddingOnly As Decimal?, weight_ChoiceWaterproof As Decimal?, weight_ChoiceWaterproof_Padded As Decimal?, weight_PremiumSyntheticLeather As Decimal?, weight_PremiumSyntheticLeather_Padded As Decimal?),
    notes As String,
    profit_Choice As Decimal?,
    profit_ChoicePadded As Decimal?,
    profit_Leather As Decimal?,
    profit_LeatherPadded As Decimal?,
    AmazonFee_Choice As Decimal?,
    AmazonFee_ChoicePadded As Decimal?,
    AmazonFee_Leather As Decimal?,
    AmazonFee_LeatherPadded As Decimal?,
    ReverbFee_Choice As Decimal?,
    ReverbFee_ChoicePadded As Decimal?,
    ReverbFee_Leather As Decimal?,
    ReverbFee_LeatherPadded As Decimal?,
    BaseCost_GrandTotal_Choice_Amazon As Decimal?,
    BaseCost_GrandTotal_ChoicePadded_Amazon As Decimal?,
    BaseCost_GrandTotal_Leather_Amazon As Decimal?,
    BaseCost_GrandTotal_LeatherPadded_Amazon As Decimal?,
    BaseCost_GrandTotal_Choice_Reverb As Decimal?,
    BaseCost_GrandTotal_ChoicePadded_Reverb As Decimal?,
    BaseCost_GrandTotal_Leather_Reverb As Decimal?,
    BaseCost_GrandTotal_LeatherPadded_Reverb As Decimal?,
    BaseCost_GrandTotal_Choice_eBay As Decimal?,
    BaseCost_GrandTotal_ChoicePadded_eBay As Decimal?,
    BaseCost_GrandTotal_Leather_eBay As Decimal?,
    BaseCost_GrandTotal_LeatherPadded_eBay As Decimal?,
    BaseCost_GrandTotal_Choice_Etsy As Decimal?,
    BaseCost_GrandTotal_ChoicePadded_Etsy As Decimal?,
    BaseCost_GrandTotal_Leather_Etsy As Decimal?,
    BaseCost_GrandTotal_LeatherPadded_Etsy As Decimal?
)
        Dim db As New DbConnectionManager()
        Dim wastePercent As Decimal = 5D

        ' Calculate shipping cost for each fabric type/combination (shipping only)
        Dim shipping_Choice As Decimal = db.GetShippingCostByWeight(If(weights.weight_ChoiceWaterproof, 0D))
        Dim shipping_ChoicePadded As Decimal = db.GetShippingCostByWeight(If(weights.weight_ChoiceWaterproof_Padded, 0D))
        Dim shipping_Leather As Decimal = db.GetShippingCostByWeight(If(weights.weight_PremiumSyntheticLeather, 0D))
        Dim shipping_LeatherPadded As Decimal = db.GetShippingCostByWeight(If(weights.weight_PremiumSyntheticLeather_Padded, 0D))

        ' Calculate base fabric cost + shipping for each type
        Dim baseFabricCost_Choice_Weight As Decimal? = If(costs.baseCost_ChoiceWaterproof, 0D) + shipping_Choice
        Dim baseFabricCost_ChoicePadding_Weight As Decimal? = If(costs.baseCost_ChoiceWaterproof_Padded, 0D) + shipping_ChoicePadded
        Dim baseFabricCost_Leather_Weight As Decimal? = If(costs.baseCost_PremiumSyntheticLeather, 0D) + shipping_Leather
        Dim baseFabricCost_LeatherPadding_Weight As Decimal? = If(costs.baseCost_PremiumSyntheticLeather_Padded, 0D) + shipping_LeatherPadded

        ' --- Labor cost calculations ---
        Dim hourlyRate As Decimal = 17D
        Dim CostLaborNoPadding As Decimal = 0.5D * hourlyRate
        Dim CostLaborWithPadding As Decimal = 1D * hourlyRate

        Dim BaseCost_Choice_Labor As Decimal? = If(baseFabricCost_Choice_Weight, 0D) + CostLaborNoPadding
        Dim BaseCost_ChoicePadding_Labor As Decimal? = If(baseFabricCost_ChoicePadding_Weight, 0D) + CostLaborWithPadding
        Dim BaseCost_Leather_Labor As Decimal? = If(baseFabricCost_Leather_Weight, 0D) + CostLaborNoPadding
        Dim BaseCost_LeatherPadding_Labor As Decimal? = If(baseFabricCost_LeatherPadding_Weight, 0D) + CostLaborWithPadding

        ' --- Marketplace Fee Calculations for eBay and Etsy ---
        Dim eBayFeePercent As Decimal = db.GetMarketplaceFeePercentage("eBay") / 100D
        Dim etsyFeePercent As Decimal = db.GetMarketplaceFeePercentage("Etsy") / 100D

        Dim eBayFee_Choice As Decimal = If(BaseCost_GrandTotal_Choice_eBay, 0D) * eBayFeePercent
        Dim eBayFee_ChoicePadded As Decimal = If(BaseCost_GrandTotal_ChoicePadded_eBay, 0D) * eBayFeePercent
        Dim eBayFee_Leather As Decimal = If(BaseCost_GrandTotal_Leather_eBay, 0D) * eBayFeePercent
        Dim eBayFee_LeatherPadded As Decimal = If(BaseCost_GrandTotal_LeatherPadded_eBay, 0D) * eBayFeePercent

        Dim EtsyFee_Choice As Decimal = If(BaseCost_GrandTotal_Choice_Etsy, 0D) * etsyFeePercent
        Dim EtsyFee_ChoicePadded As Decimal = If(BaseCost_GrandTotal_ChoicePadded_Etsy, 0D) * etsyFeePercent
        Dim EtsyFee_Leather As Decimal = If(BaseCost_GrandTotal_Leather_Etsy, 0D) * etsyFeePercent
        Dim EtsyFee_LeatherPadded As Decimal = If(BaseCost_GrandTotal_LeatherPadded_Etsy, 0D) * etsyFeePercent

        ' --- Retail Price Placeholders (replace with your actual calculations) ---
        Dim RetailPrice_Choice_Amazon As Decimal = 0D
        Dim RetailPrice_ChoicePadded_Amazon As Decimal = 0D
        Dim RetailPrice_Leather_Amazon As Decimal = 0D
        Dim RetailPrice_LeatherPadded_Amazon As Decimal = 0D
        Dim RetailPrice_Choice_Reverb As Decimal = 0D
        Dim RetailPrice_ChoicePadded_Reverb As Decimal = 0D
        Dim RetailPrice_Leather_Reverb As Decimal = 0D
        Dim RetailPrice_LeatherPadded_Reverb As Decimal = 0D
        Dim RetailPrice_Choice_eBay As Decimal = 0D
        Dim RetailPrice_ChoicePadded_eBay As Decimal = 0D
        Dim RetailPrice_Leather_eBay As Decimal = 0D
        Dim RetailPrice_LeatherPadded_eBay As Decimal = 0D
        Dim RetailPrice_Choice_Etsy As Decimal = 0D
        Dim RetailPrice_ChoicePadded_Etsy As Decimal = 0D
        Dim RetailPrice_Leather_Etsy As Decimal = 0D
        Dim RetailPrice_LeatherPadded_Etsy As Decimal = 0D

        ' Insert into history table (all arguments, in order)
        db.InsertModelHistoryCostRetailPricing(
        modelId,
        costs.costPerSqInch_ChoiceWaterproof,
        costs.costPerSqInch_PremiumSyntheticLeather,
        costs.costPerSqInch_Padding,
        totalFabricSquareInches,
        wastePercent,
        costs.baseCost_ChoiceWaterproof,
        costs.baseCost_PremiumSyntheticLeather,
        costs.baseCost_ChoiceWaterproof_Padded,
        costs.baseCost_PremiumSyntheticLeather_Padded,
        costs.baseCost_PaddingOnly,
        weights.weight_PaddingOnly,
        weights.weight_ChoiceWaterproof,
        weights.weight_ChoiceWaterproof_Padded,
        weights.weight_PremiumSyntheticLeather,
        weights.weight_PremiumSyntheticLeather_Padded,
        shipping_Choice,
        shipping_ChoicePadded,
        shipping_Leather,
        shipping_LeatherPadded,
        baseFabricCost_Choice_Weight,
        baseFabricCost_ChoicePadding_Weight,
        baseFabricCost_Leather_Weight,
        baseFabricCost_LeatherPadding_Weight,
        BaseCost_Choice_Labor,
        BaseCost_ChoicePadding_Labor,
        BaseCost_Leather_Labor,
        BaseCost_LeatherPadding_Labor,
        If(profit_Choice, 0D),
        If(profit_ChoicePadded, 0D),
        If(profit_Leather, 0D),
        If(profit_LeatherPadded, 0D),
        If(AmazonFee_Choice, 0D),
        If(AmazonFee_ChoicePadded, 0D),
        If(AmazonFee_Leather, 0D),
        If(AmazonFee_LeatherPadded, 0D),
        If(ReverbFee_Choice, 0D),
        If(ReverbFee_ChoicePadded, 0D),
        If(ReverbFee_Leather, 0D),
        If(ReverbFee_LeatherPadded, 0D),
        eBayFee_Choice,
        eBayFee_ChoicePadded,
        eBayFee_Leather,
        eBayFee_LeatherPadded,
        EtsyFee_Choice,
        EtsyFee_ChoicePadded,
        EtsyFee_Leather,
        EtsyFee_LeatherPadded,
        If(BaseCost_GrandTotal_Choice_Amazon, 0D),
        If(BaseCost_GrandTotal_ChoicePadded_Amazon, 0D),
        If(BaseCost_GrandTotal_Leather_Amazon, 0D),
        If(BaseCost_GrandTotal_LeatherPadded_Amazon, 0D),
        If(BaseCost_GrandTotal_Choice_Reverb, 0D),
        If(BaseCost_GrandTotal_ChoicePadded_Reverb, 0D),
        If(BaseCost_GrandTotal_Leather_Reverb, 0D),
        If(BaseCost_GrandTotal_LeatherPadded_Reverb, 0D),
        If(BaseCost_GrandTotal_Choice_eBay, 0D),
        If(BaseCost_GrandTotal_ChoicePadded_eBay, 0D),
        If(BaseCost_GrandTotal_Leather_eBay, 0D),
        If(BaseCost_GrandTotal_LeatherPadded_eBay, 0D),
        If(BaseCost_GrandTotal_Choice_Etsy, 0D),
        If(BaseCost_GrandTotal_ChoicePadded_Etsy, 0D),
        If(BaseCost_GrandTotal_Leather_Etsy, 0D),
        If(BaseCost_GrandTotal_LeatherPadded_Etsy, 0D),
        RetailPrice_Choice_Amazon,
        RetailPrice_ChoicePadded_Amazon,
        RetailPrice_Leather_Amazon,
        RetailPrice_LeatherPadded_Amazon,
        RetailPrice_Choice_Reverb,
        RetailPrice_ChoicePadded_Reverb,
        RetailPrice_Leather_Reverb,
        RetailPrice_LeatherPadded_Reverb,
        RetailPrice_Choice_eBay,
        RetailPrice_ChoicePadded_eBay,
        RetailPrice_Leather_eBay,
        RetailPrice_LeatherPadded_eBay,
        RetailPrice_Choice_Etsy,
        RetailPrice_ChoicePadded_Etsy,
        RetailPrice_Leather_Etsy,
        RetailPrice_LeatherPadded_Etsy,
        notes
    )
    End Sub
    ' ================================================================
    ' File: HelperFunctions.vb
    ' Purpose:
    '   Utility helpers for interacting with Windows Task Scheduler and
    '   waiting for a file (DailySchema.json) to be created/updated.
    '
    ' Public Functions:
    '   - RunScheduledTask(taskNameOrPath As String) As Boolean
    '   - TaskExists(taskNameOrPath As String) As Boolean
    '   - GetTaskLastRunTime(taskNameOrPath As String) As DateTime?
    '   - WaitForFileChange(path As String, previousWriteTimeUtc As DateTime?, timeoutSeconds As Integer) As Boolean
    '
    ' Dependencies:
    '   - schtasks.exe (built into Windows)
    '   - No external NuGet packages required
    ' ================================================================

    ' ------------------------------------------------------------
    ' Function: RunScheduledTask
    ' Purpose : Start a scheduled task immediately using schtasks.exe
    ' Params  : taskNameOrPath - exact task name or full path (\Folder\TaskName)
    ' Returns : True if task started successfully, False otherwise
    ' ------------------------------------------------------------
    Public Function RunScheduledTask(taskNameOrPath As String) As Boolean
        If String.IsNullOrWhiteSpace(taskNameOrPath) Then Throw New ArgumentException("Task name/path is required.", NameOf(taskNameOrPath))

        Dim psi As New ProcessStartInfo() With {
            .FileName = "schtasks.exe",
            .Arguments = "/run /tn " & QuoteArg(taskNameOrPath),
            .CreateNoWindow = True,
            .UseShellExecute = False,
            .RedirectStandardOutput = True,
            .RedirectStandardError = True
        }

        Using p As Process = Process.Start(psi)
            p.WaitForExit()
            Return p.ExitCode = 0
        End Using
    End Function


    ' ------------------------------------------------------------
    ' Function: TaskExists
    ' What it does:
    '   Checks if a scheduled task exists using:
    '     schtasks.exe /query /tn "<TaskName>"
    ' Parameters:
    '   taskName : The Task Scheduler "Task Name"
    ' Returns:
    '   Boolean  : True if the task is found (exit code 0)
    ' ------------------------------------------------------------
    Public Function TaskExists(taskNameOrPath As String) As Boolean
        If String.IsNullOrWhiteSpace(taskNameOrPath) Then Return False

        Dim psi As New ProcessStartInfo() With {
            .FileName = "schtasks.exe",
            .Arguments = "/query /tn " & QuoteArg(taskNameOrPath),
            .CreateNoWindow = True,
            .UseShellExecute = False,
            .RedirectStandardOutput = True,
            .RedirectStandardError = True
        }

        Using p As Process = Process.Start(psi)
            p.WaitForExit()
            Return p.ExitCode = 0
        End Using
    End Function

    ' ------------------------------------------------------------
    ' Function: GetTaskLastRunTime
    ' What it does:
    '   Queries verbose task info and attempts to parse "Last Run Time".
    '   Uses:
    '     schtasks.exe /query /tn "<TaskName>" /fo LIST /v
    ' Parameters:
    '   taskName : The Task Scheduler "Task Name"
    ' Returns:
    '   Nullable(Of DateTime) : Last run time if parse succeeds; otherwise Nothing
    ' ------------------------------------------------------------
    Public Function GetTaskLastRunTime(taskNameOrPath As String) As DateTime?
        If String.IsNullOrWhiteSpace(taskNameOrPath) Then Return Nothing

        Dim psi As New ProcessStartInfo() With {
            .FileName = "schtasks.exe",
            .Arguments = "/query /tn " & QuoteArg(taskNameOrPath) & " /fo LIST /v",
            .CreateNoWindow = True,
            .UseShellExecute = False,
            .RedirectStandardOutput = True,
            .RedirectStandardError = True,
            .StandardOutputEncoding = Encoding.UTF8
        }

        Using p As Process = Process.Start(psi)
            Dim output As String = p.StandardOutput.ReadToEnd()
            p.WaitForExit()

            If p.ExitCode <> 0 OrElse String.IsNullOrEmpty(output) Then Return Nothing

            Dim m As Match = Regex.Match(output, "Last Run Time\s*:\s*(.+)", RegexOptions.IgnoreCase)
            If m.Success Then
                Dim raw As String = m.Groups(1).Value.Trim()
                Dim dt As DateTime
                If DateTime.TryParse(raw, dt) Then Return dt
            End If

            Return Nothing
        End Using
    End Function


    ' ------------------------------------------------------------
    ' Function: WaitForFileChange
    ' Purpose : Wait until a file is created or modified
    ' Params  : path - file path to watch
    '           previousWriteTimeUtc - last known write time (Nothing if unknown)
    '           timeoutSeconds - how long to wait
    ' Returns : True if file updated before timeout, False otherwise
    ' ------------------------------------------------------------
    Public Function WaitForFileChange(path As String, previousWriteTimeUtc As DateTime?, timeoutSeconds As Integer) As Boolean
        Dim stopAt As DateTime = DateTime.UtcNow.AddSeconds(Math.Max(1, timeoutSeconds))

        Do
            If File.Exists(path) Then
                Dim current As DateTime = File.GetLastWriteTimeUtc(path)
                If Not previousWriteTimeUtc.HasValue OrElse current > previousWriteTimeUtc.Value Then
                    Return True
                End If
            End If

            System.Threading.Thread.Sleep(500) ' poll every 0.5s
        Loop While DateTime.UtcNow < stopAt

        Return False
    End Function
    ' ------------------------------------------------------------
    ' Function: QuoteArg
    ' What it does:
    '   Safely quotes an argument for the command line.
    ' Parameters:
    '   s : raw string to quote
    ' Returns:
    '   String : quoted argument (e.g., "My Task Name")
    ' ------------------------------------------------------------
    Private Function QuoteArg(s As String) As String
        If s.Contains("""") Then Return s
        Return """" & s & """"
    End Function











    ' ================================================================
    ' File/Module: HelperFunctions.vb
    ' Purpose   : Utility helpers for building GitHub file links from a
    '             local repository path and displaying them in the UI.
    '             Creates links like:
    '             https://github.com/{owner}/{repo}/blob/{branch}/{path}
    ' Dependencies:
    '   - No external NuGet packages required
    ' Notes:
    '   - Excludes common build/IDE folders (.git, .vs, bin, obj) by default
    '   - Uses Uri.EscapeDataString to safely encode URL path segments
    ' ================================================================

    ' ------------------------------------------------------------
    ' Function: BuildGitHubLinks
    ' Purpose : Enumerate files under repoRoot and return (LocalPath, Url)
    ' Params  : repoRoot   - local repository root containing GGC.sln
    '           owner      - GitHub account/organization (e.g., "ksmartz")
    '           repo       - GitHub repo name (e.g., "GigGearCovers")
    '           branch     - branch name ("main" or "master")
    '           includeExt - list of file extensions to include (".vb", ".resx", ".sln", ".vbproj", etc.)
    '           extraExcl  - additional relative folders to exclude (optional)
    ' Returns : List(Of Tuple(Of String, String)) -> (localPath, githubUrl)
    ' Depends : None
    ' ------------------------------------------------------------
    Public Function BuildGitHubLinks(repoRoot As String,
                                     owner As String,
                                     repo As String,
                                     branch As String,
                                     includeExt As IEnumerable(Of String),
                                     Optional extraExcl As IEnumerable(Of String) = Nothing) As List(Of Tuple(Of String, String))

        If String.IsNullOrWhiteSpace(repoRoot) OrElse Not Directory.Exists(repoRoot) Then
            Throw New DirectoryNotFoundException("Repository root not found: " & repoRoot)
        End If

        Dim exts As HashSet(Of String) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each e In includeExt
            Dim norm = If(e.StartsWith("."c), e, "." & e)
            exts.Add(norm)
        Next

        Dim defaultExclude As String() = {".git", ".vs", "bin", "obj"}
        Dim exclude As HashSet(Of String) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each d In defaultExclude
            exclude.Add(d.TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar))
        Next
        If extraExcl IsNot Nothing Then
            For Each d In extraExcl
                exclude.Add(d.TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar))
            Next
        End If

        Dim results As New List(Of Tuple(Of String, String))()

        ' Enumerate all files and filter by extension & exclude dirs
        Dim allFiles = Directory.EnumerateFiles(repoRoot, "*.*", SearchOption.AllDirectories)

        For Each full In allFiles
            Dim rel As String = GetRelativePath(repoRoot, full) ' e.g., "Forms\MyForm.vb"

            ' Skip excluded folders if the relative path starts with them
            Dim skip As Boolean = False
            For Each ex In exclude
                If rel.StartsWith(ex & Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(rel, ex, StringComparison.OrdinalIgnoreCase) Then
                    skip = True
                    Exit For
                End If
            Next
            If skip Then Continue For

            ' Filter by extension
            Dim ext As String = Path.GetExtension(full)
            If Not exts.Contains(ext) Then Continue For

            ' Build GitHub URL
            Dim ghUrl As String = BuildGitHubBlobUrl(owner, repo, branch, rel)

            results.Add(Tuple.Create(rel, ghUrl))
        Next

        ' Sort for convenience (by relative path)
        results.Sort(Function(a, b) StringComparer.OrdinalIgnoreCase.Compare(a.Item1, b.Item1))
        Return results
    End Function

    ' ------------------------------------------------------------
    ' Function: BuildGitHubBlobUrl
    ' Purpose : Create a GitHub blob URL for a relative path
    ' Params  : owner/repo/branch - GitHub coordinates
    '           relativePath      - Windows-style relative path (uses \)
    ' Returns : Full URL string
    ' ------------------------------------------------------------
    Public Function BuildGitHubBlobUrl(owner As String, repo As String, branch As String, relativePath As String) As String
        ' Convert to URL path with forward slashes and encode each segment
        Dim parts = relativePath.Split(New Char() {Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar}, StringSplitOptions.RemoveEmptyEntries)
        Dim encodedSegments As New List(Of String)
        For Each seg In parts
            encodedSegments.Add(Uri.EscapeDataString(seg))
        Next
        Dim encodedPath As String = String.Join("/", encodedSegments)
        Return $"https://github.com/{owner}/{repo}/blob/{Uri.EscapeDataString(branch)}/{encodedPath}"
    End Function

    ' ------------------------------------------------------------
    ' Function: GetRelativePath
    ' Purpose : Return a Windows relative path from baseDir to fullPath
    ' ------------------------------------------------------------
    Public Function GetRelativePath(baseDir As String, fullPath As String) As String
        Dim baseUri As New Uri(EnsureTrailingSlash(New Uri(Path.GetFullPath(baseDir)).ToString()))
        Dim fileUri As New Uri(Path.GetFullPath(fullPath))
        Dim relUri As Uri = baseUri.MakeRelativeUri(fileUri)
        Dim rel As String = Uri.UnescapeDataString(relUri.ToString()).Replace("/"c, Path.DirectorySeparatorChar)
        Return rel
    End Function

    Private Function EnsureTrailingSlash(s As String) As String
        If s.EndsWith("/") Then Return s
        Return s & "/"
    End Function

End Module

