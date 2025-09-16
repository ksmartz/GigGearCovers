Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text.Json
Imports System.Threading.Tasks
Imports System.Text    ' <- needed for StringBuilder


' Result types you can use anywhere
Public Class VariationOutcome
    Public Property Sku As String
    Public Property VariationId As Integer
End Class

Public Class PublishResult
    Public Property Success As Boolean
    Public Property Message As String
    Public Property ProductId As Integer
    Public Property ParentSku As String
    Public Property VariationResults As List(Of VariationOutcome) = New List(Of VariationOutcome)()
End Class

' All shared/static (no instances of this type needed)
Public Module WooVariationPublisher

    Private ReadOnly JsonOptions As New JsonSerializerOptions With {
        .PropertyNameCaseInsensitive = True
    }

    Private Function Sanitize(input As String) As String
        If input Is Nothing Then Return ""
        Dim sb As New StringBuilder(input.Length)
        For Each ch In input
            If Char.IsLetterOrDigit(ch) Then
                sb.Append(Char.ToUpperInvariant(ch))
            ElseIf Char.IsWhiteSpace(ch) OrElse ch = "-"c OrElse ch = "_"c Then
                sb.Append("-"c)
            End If
        Next
        ' collapse doubles like "--"
        Dim s = sb.ToString()
        While s.Contains("--")
            s = s.Replace("--", "-")
        End While
        Return s.Trim("-"c)
    End Function

    Private Function ExtractIdFromJson(json As String) As Integer
        If String.IsNullOrWhiteSpace(json) Then Return 0
        Try
            Using doc = JsonDocument.Parse(json)
                Dim root = doc.RootElement
                If root.TryGetProperty("id", Nothing) Then
                    Return root.GetProperty("id").GetInt32()
                End If
            End Using
        Catch
            ' ignore parse errors, return 0
        End Try
        Return 0
    End Function

    '==============================================================================
    ' Function: PublishModelAsync
    ' Class/Module: WooVariationPublisher
    ' Purpose :
    '   Publish a model to WooCommerce as a VARIABLE parent with Fabric × Color
    '   variations (using DB ABBREVIATIONS). Logs requests/responses and maps IDs.
    '
    ' Depends on:
    '   DbConnectionManager:
    '       - GetModelRow(modelId) As DataRow
    '       - GetEquipmentTypeForSeries(seriesId) -> (Id As Integer?, Name As String)
    '       - GetFabricAbbreviationMap() -> Dictionary(Of String,String) ' name -> 1-char
    '       - GetColorAbbreviationMap()  -> Dictionary(Of String,String) ' name -> 3-char
    '       - UpsertMpWooProduct(...), UpsertMpWooVariation(...)
    '   WooCommerceAPI: Create/Update product/variation + GET by SKU helpers
    '   Helpers present in this module: ResolveSeriesId, ResolveNames, SafeRowString,
    '       FindProductIdBySku, PollFindProductIdBySku, FindVariationIdBySku,
    '       PollFindVariationIdBySku, GetWooErrorText, ExtractIdFromJson, JsonOptions
    '   ModelSkuBuilder.GenerateChildSkuFromParentCodes(parentSku, fCode, cCode)
    '------------------------------------------------------------------------------
    Public Async Function PublishModelAsync(modelId As Integer) As Task(Of PublishResult)
        Dim res As New PublishResult()

        Try
            ' 1) Load model basics from DB
            Dim m As DataRow = DbConnectionManager.GetModelRow(modelId)
            If m Is Nothing Then
                res.Success = False
                res.Message = $"Model {modelId} not found."
                Return res
            End If

            Dim modelName As String = SafeRowString(m, "ModelName")
            Dim parentSku As String = SafeRowString(m, "ParentSKU")
            If String.IsNullOrWhiteSpace(parentSku) Then
                res.Success = False
                res.Message = "ParentSKU missing. Please set ParentSKU for the model first."
                Return res
            End If

            ' Resolve series/category + names from DB (UI-agnostic)
            Dim seriesId As Integer = ResolveSeriesId(m, modelId)
            Dim names = ResolveNames(modelId)
            Dim manufacturerName As String = names.ManufacturerName
            Dim seriesName As String = names.SeriesName

            Dim eq = DbConnectionManager.GetEquipmentTypeForSeries(seriesId) ' (Id As Integer?, Name As String)
            Dim wooCategoryId As Integer = If(eq.Id.HasValue, eq.Id.Value, 0)

            ' 2) Pull ABBREVIATION MAPS (name -> code) and build attribute option lists as CODES
            Dim fabricMap As Dictionary(Of String, String) = DbConnectionManager.GetFabricAbbreviationMap() ' 1-char
            Dim colorMap As Dictionary(Of String, String) = DbConnectionManager.GetColorAbbreviationMap()  ' 3-char

            Dim fabricCodes As String() =
            fabricMap.Values.
                Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                Select(Function(s) s.Trim().ToUpperInvariant()).
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToArray()

            Dim colorCodes As String() =
            colorMap.Values.
                Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                Select(Function(s) s.Trim().ToUpperInvariant()).
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToArray()

            If fabricCodes.Length = 0 Then
                res.Success = False : res.Message = "No Fabric abbreviations found." : Return res
            End If
            If colorCodes.Length = 0 Then
                res.Success = False : res.Message = "No Color abbreviations found." : Return res
            End If

            ' Parent attributes must list the exact options used by the variations
            Dim attributes = New List(Of Object) From {
            New With {.name = "Fabric", .variation = True, .visible = True, .options = fabricCodes},
            New With {.name = "Color", .variation = True, .visible = True, .options = colorCodes}
        }

            Dim title As String = $"{manufacturerName} {seriesName} {modelName}".Trim()
            Dim productPayload = New With {
            .name = title,
            .type = "variable",
            .sku = parentSku,
            .categories = If(wooCategoryId > 0, New Object() {New With {.id = wooCategoryId}}, New Object() {}),
            .attributes = attributes
        }

            ' 3) UPSERT PARENT by SKU
            Dim parentId As Integer = Await FindProductIdBySku(parentSku)
            Dim productReq As String = JsonSerializer.Serialize(productPayload, JsonOptions)

            If parentId > 0 Then
                Dim upd = Await WooCommerceAPI.UpdateProductAsync(parentId, productPayload)
                If upd.Status >= 200 AndAlso upd.Status < 300 Then
                    DbConnectionManager.InsertMpWooSyncLog("update_product_parent", parentSku, parentId, True, upd.Status.ToString(), productReq, upd.Body, Nothing)
                Else
                    Dim errText = GetWooErrorText(upd.Body)
                    DbConnectionManager.InsertMpWooSyncLog("update_product_parent", parentSku, parentId, False, upd.Status.ToString(), productReq, upd.Body, errText)
                    res.Success = False : res.Message = $"Update product failed: {errText} (HTTP {upd.Status})"
                    Return res
                End If
            Else
                Dim crt = Await WooCommerceAPI.CreateProductAsync(productPayload)
                If crt.Status >= 200 AndAlso crt.Status < 300 Then
                    parentId = ExtractIdFromJson(crt.Body)
                    DbConnectionManager.InsertMpWooSyncLog("create_product_parent", parentSku, parentId, True, crt.Status.ToString(), productReq, crt.Body, Nothing)
                Else
                    Dim errText = GetWooErrorText(crt.Body)
                    If errText.IndexOf("under processing", StringComparison.OrdinalIgnoreCase) >= 0 _
                   OrElse errText.IndexOf("already exists", StringComparison.OrdinalIgnoreCase) >= 0 Then

                        parentId = Await PollFindProductIdBySku(parentSku, maxAttempts:=30, delayMs:=2000)

                        If parentId = 0 Then
                            Await Task.Delay(3000)
                            Dim retry = Await WooCommerceAPI.CreateProductAsync(productPayload)
                            If retry.Status >= 200 AndAlso retry.Status < 300 Then
                                parentId = ExtractIdFromJson(retry.Body)
                                DbConnectionManager.InsertMpWooSyncLog("create_product_parent_retry", parentSku, parentId, True, retry.Status.ToString(), productReq, retry.Body, "Recovered after retry")
                            Else
                                Dim retryErr = GetWooErrorText(retry.Body)
                                DbConnectionManager.InsertMpWooSyncLog("create_product_parent", parentSku, Nothing, False, retry.Status.ToString(), productReq, retry.Body, $"Retry failed: {retryErr}")
                                res.Success = False
                                res.Message = $"Create product failed: {retryErr} (HTTP {retry.Status})"
                                Return res
                            End If
                        Else
                            DbConnectionManager.InsertMpWooSyncLog("create_product_parent", parentSku, parentId, True, crt.Status.ToString(), productReq, crt.Body, "Recovered via SKU lookup after processing")
                        End If
                    Else
                        DbConnectionManager.InsertMpWooSyncLog("create_product_parent", parentSku, Nothing, False, crt.Status.ToString(), productReq, crt.Body, errText)
                        res.Success = False
                        res.Message = $"Create product failed: {errText} (HTTP {crt.Status})"
                        Return res
                    End If
                End If
            End If

            ' Mirror to mapping table
            Dim mpWooProductId As Integer =
            DbConnectionManager.UpsertMpWooProduct(modelId, parentSku, parentId, wooCategoryId, "OK", "Product upserted")

            ' 4) UPSERT VARIATIONS (Fabric × Color) — iterate over CODES
            ' === Variations: iterate over the CODES ===
            For Each fCode In fabricCodes
                For Each cCode In colorCodes
                    ' Build from parent with reserved tail removed
                    Dim childSku As String = ModelSkuBuilder.GenerateChildSkuFromParentCodes(parentSku, fCode, cCode)
                    childSku = If(childSku, "").Trim()
                    If childSku.Length = 0 Then
                        ' Log and skip — do NOT touch DB with a blank key
                        DbConnectionManager.InsertMpWooSyncLog("skip_variation_child", "(empty-childsku)", Nothing, False, "0",
                                                   $"parent={parentSku}; f={fCode}; c={cCode}", "", "ChildSKU blank; skipped")
                        Continue For
                    End If

                    Dim variationPayload = New With {
            .sku = childSku,
            .attributes = New Object() {
                New With {.name = "Fabric", .option = fCode},
                New With {.name = "Color", .option = cCode}
            }
        }

                    Dim vReq As String = JsonSerializer.Serialize(variationPayload, JsonOptions)

                    Dim varId As Integer = Await FindVariationIdBySku(parentId, childSku)
                    If varId > 0 Then
                        Dim upd = Await WooCommerceAPI.UpdateVariationAsync(parentId, varId, variationPayload)
                        If upd.Status >= 200 AndAlso upd.Status < 300 Then
                            DbConnectionManager.InsertMpWooSyncLog("update_variation_child", childSku, varId, True, upd.Status.ToString(), vReq, upd.Body, Nothing)
                            DbConnectionManager.UpsertMpWooVariation(mpWooProductId, childSku, varId, fCode, cCode, "OK", "Variation upserted")
                            res.VariationResults.Add(New VariationOutcome With {.Sku = childSku, .VariationId = varId})
                        Else
                            Dim errText = GetWooErrorText(upd.Body)
                            DbConnectionManager.InsertMpWooSyncLog("update_variation_child", childSku, varId, False, upd.Status.ToString(), vReq, upd.Body, errText)
                        End If
                    Else
                        Dim crtV = Await WooCommerceAPI.CreateVariationAsync(parentId, variationPayload)
                        If crtV.Status >= 200 AndAlso crtV.Status < 300 Then
                            Dim newId As Integer = ExtractIdFromJson(crtV.Body)
                            DbConnectionManager.InsertMpWooSyncLog("create_variation_child", childSku, newId, True, crtV.Status.ToString(), vReq, crtV.Body, Nothing)
                            DbConnectionManager.UpsertMpWooVariation(mpWooProductId, childSku, newId, fCode, cCode, "OK", "Variation upserted")
                            res.VariationResults.Add(New VariationOutcome With {.Sku = childSku, .VariationId = newId})
                        Else
                            Dim errText = GetWooErrorText(crtV.Body)
                            DbConnectionManager.InsertMpWooSyncLog("create_variation_child", childSku, Nothing, False, crtV.Status.ToString(), vReq, crtV.Body, errText)
                        End If
                    End If
                Next
            Next


            res.Success = True
            res.ProductId = parentId
            res.ParentSku = parentSku
            res.Message = $"Product {parentId} with {res.VariationResults.Count} variations published."
            Return res

        Catch ex As Exception
            res.Success = False
            res.Message = $"Publish failed for Model {modelId}: {ex.Message}"
            Return res
        End Try
    End Function


    ' --- Helpers used by PublishModelAsync (paste into WooVariationPublisher) ---

    Private Async Function FindProductIdBySku(sku As String) As Task(Of Integer)
        If String.IsNullOrWhiteSpace(sku) Then Return 0
        Dim r = Await WooCommerceAPI.GetProductsBySkuAsync(sku)
        If r.Status >= 200 AndAlso r.Status < 300 Then
            Return ParseFirstIdFromArrayJson(r.Body)
        End If
        Return 0
    End Function

    Private Async Function PollFindProductIdBySku(sku As String, maxAttempts As Integer, delayMs As Integer) As Task(Of Integer)
        Dim attempts = If(maxAttempts <= 0, 30, maxAttempts)      ' ← up to ~60s
        Dim pause = If(delayMs <= 0, 2000, delayMs)
        For i = 1 To attempts
            Dim id = Await FindProductIdBySku(sku)
            If id > 0 Then Return id
            Await Task.Delay(pause)
        Next
        Return 0
    End Function


    Private Async Function FindVariationIdBySku(parentId As Integer, sku As String) As Task(Of Integer)
        If parentId <= 0 OrElse String.IsNullOrWhiteSpace(sku) Then Return 0
        Dim r = Await WooCommerceAPI.GetVariationsBySkuAsync(parentId, sku)
        If r.Status >= 200 AndAlso r.Status < 300 Then
            Return ParseFirstIdFromArrayJson(r.Body)
        End If
        Return 0
    End Function

    Private Async Function PollFindVariationIdBySku(parentId As Integer, sku As String, maxAttempts As Integer, delayMs As Integer) As Task(Of Integer)
        Dim attempts = Math.Max(1, maxAttempts)
        Dim pause = Math.Max(50, delayMs)
        For i = 1 To attempts
            Dim id = Await FindVariationIdBySku(parentId, sku)
            If id > 0 Then Return id
            Await Task.Delay(pause)
        Next
        Return 0
    End Function

    ' Parses: [ { "id": 123, ... }, ... ]
    Private Function ParseFirstIdFromArrayJson(body As String) As Integer
        If String.IsNullOrWhiteSpace(body) Then Return 0
        Try
            Using doc = JsonDocument.Parse(body)
                Dim root = doc.RootElement
                If root.ValueKind = JsonValueKind.Array AndAlso root.GetArrayLength() > 0 Then
                    Dim first = root(0)
                    Dim idProp As System.Text.Json.JsonElement
                    If first.TryGetProperty("id", idProp) Then
                        Dim id As Integer = 0
                        Integer.TryParse(idProp.ToString(), id)
                        Return id
                    End If
                End If
            End Using
        Catch
            ' swallow parse errors; treat as not found
        End Try
        Return 0
    End Function


    '------------------------------------------------------------------------------
    ' Helper: SafeRowString
    '------------------------------------------------------------------------------
    Private Function SafeRowString(r As DataRow, colName As String) As String
        If r Is Nothing OrElse r.Table Is Nothing OrElse Not r.Table.Columns.Contains(colName) Then Return ""
        Dim v = r(colName)
        If v Is Nothing OrElse v Is DBNull.Value Then Return ""
        Return v.ToString().Trim()
    End Function

    '------------------------------------------------------------------------------
    ' Helper: GetWooErrorText
    ' Purpose: Pull a friendly message from Woo's error JSON; fallback to raw body.
    '------------------------------------------------------------------------------
    Private Function GetWooErrorText(body As String) As String
        If String.IsNullOrWhiteSpace(body) Then Return "Bad request"
        Try
            ' Woo sends: { "code": "...", "message": "reason", "data": { "status": 400 } }
            Dim doc = JsonDocument.Parse(body)
            Dim root = doc.RootElement
            If root.TryGetProperty("message", Nothing) Then
                Return root.GetProperty("message").GetString()
            End If
        Catch
            ' ignore parse errors; fall back to body
        End Try
        ' Trim very long responses
        If body.Length > 400 Then Return body.Substring(0, 400) & "..."
        Return body
    End Function


    '==============================================================================
    ' Function: ResolveNames
    ' Class/Module: WooVariationPublisher
    ' Purpose :
    '   Return ManufacturerName and SeriesName from DB by modelId (UI-agnostic).
    '
    ' Change (2025-09-15):
    '   - Join Manufacturer via Series (s.FK_ManufacturerId) instead of Model.
    '
    ' Dependencies:
    '   - DbConnectionManager.GetConnection() As SqlConnection
    '   - DbConnectionManager.EnsureOpen(conn As SqlConnection)
    '
    ' Tables (adjust names if your actual schema differs):
    '   - Model (PK_ModelId, FK_SeriesId)
    '   - ModelSeries (PK_SeriesId, FK_ManufacturerId, SeriesName)
    '   - ModelManufacturers (PK_ManufacturerId, ManufacturerName)
    '
    ' Date: 2025-09-15
    '------------------------------------------------------------------------------
    Private Function ResolveNames(modelId As Integer) As (ManufacturerName As String, SeriesName As String)
        Dim manu As String = ""
        Dim series As String = ""

        Using conn As SqlConnection = DbConnectionManager.GetConnection()
            DbConnectionManager.EnsureOpen(conn)
            Using cmd As New SqlCommand("
            SELECT TOP (1)
                   mf.ManufacturerName,
                   s.SeriesName
            FROM Model              AS m
            INNER JOIN ModelSeries  AS s  ON s.PK_SeriesId        = m.FK_SeriesId
            INNER JOIN ModelManufacturers AS mf ON mf.PK_ManufacturerId = s.FK_ManufacturerId
            WHERE m.PK_ModelId = @mid;", conn)

                cmd.Parameters.Add("@mid", SqlDbType.Int).Value = modelId

                Using rd = cmd.ExecuteReader()
                    If rd.Read() Then
                        manu = If(TryCast(rd("ManufacturerName"), String), Nothing)
                        series = If(TryCast(rd("SeriesName"), String), Nothing)
                    End If
                End Using
            End Using
        End Using

        manu = If(String.IsNullOrWhiteSpace(manu), "", manu.Trim())
        series = If(String.IsNullOrWhiteSpace(series), "", series.Trim())
        Return (manu, series)
    End Function

    '==============================================================================
    ' Function: ResolveSeriesId
    ' Class/Module: WooVariationPublisher
    ' Purpose :
    '   Return SeriesId for a model regardless of DataRow shape. Tries common column
    '   names; if not present/valid, queries DB by ModelId.
    '
    ' Dependencies:
    '   DbConnectionManager.GetConnection(), DbConnectionManager.EnsureOpen(conn)
    '
    ' Date: 2025-09-15
    '------------------------------------------------------------------------------
    '------------------------------------------------------------------------------
    ' Helper: ResolveSeriesId  (robust to varying DataRow shapes; falls back to DB)
    '------------------------------------------------------------------------------
    Private Function ResolveSeriesId(modelRow As DataRow, modelId As Integer) As Integer
        If modelRow IsNot Nothing AndAlso modelRow.Table IsNot Nothing Then
            Dim t = modelRow.Table
            Dim tryCols() As String = {"PK_SeriesID", "PK_SeriesId", "FK_SeriesId", "SeriesId"}
            For Each col In tryCols
                If t.Columns.Contains(col) Then
                    Dim v = modelRow(col)
                    Dim sid As Integer
                    If v IsNot Nothing AndAlso v IsNot DBNull.Value AndAlso Integer.TryParse(v.ToString(), sid) AndAlso sid > 0 Then
                        Return sid
                    End If
                End If
            Next
        End If

        Using conn As SqlConnection = DbConnectionManager.GetConnection()
            DbConnectionManager.EnsureOpen(conn)
            Using cmd As New SqlCommand("
            SELECT TOP (1) s.PK_SeriesId
            FROM Model AS m
            INNER JOIN ModelSeries AS s ON s.PK_SeriesId = m.FK_SeriesId
            WHERE m.PK_ModelId = @mid;", conn)

                cmd.Parameters.Add("@mid", SqlDbType.Int).Value = modelId
                Dim o = cmd.ExecuteScalar()
                If o Is Nothing OrElse o Is DBNull.Value Then
                    Throw New InvalidOperationException($"Could not resolve SeriesId for ModelId {modelId}.")
                End If
                Dim sid As Integer
                If Not Integer.TryParse(o.ToString(), sid) OrElse sid <= 0 Then
                    Throw New InvalidOperationException($"SeriesId lookup returned invalid value for ModelId {modelId}.")
                End If
                Return sid
            End Using
        End Using
    End Function


End Module


