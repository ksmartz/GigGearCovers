Option Strict On
Option Explicit On

Imports System.Text
Imports System.Text.Json

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

    ' MAIN ENTRY POINT
    Public Async Function PublishModelAsync(modelId As Integer) As Threading.Tasks.Task(Of PublishResult)
        Dim res As New PublishResult()

        ' 1) Load model basics (this method must exist in DbConnectionManager as Shared)
        Dim m = DbConnectionManager.GetModelRow(modelId)
        If m Is Nothing Then
            res.Success = False
            res.Message = $"Model {modelId} not found."
            Return res
        End If

        Dim modelName As String = CStr(m("ModelName"))
        Dim parentSku As String = If(TryCast(m("ParentSKU"), String), Nothing)
        Dim seriesId As Integer = CInt(m("PK_SeriesId"))
        Dim seriesName As String = CStr(m("SeriesName"))
        Dim manufacturerName As String = CStr(m("ManufacturerName"))

        If String.IsNullOrWhiteSpace(parentSku) Then
            res.Success = False
            res.Message = "ParentSKU missing. Please set ParentSKU for the model first."
            Return res
        End If

        ' 2) Category (use your own resolver if different)
        Dim eq = DbConnectionManager.GetEquipmentTypeForSeries(seriesId)
        Dim wooCategoryId As Integer = If(eq.Id.HasValue, eq.Id.Value, 0)

        ' 3) Variation attribute lists (Shared helpers in DbConnectionManager)
        Dim fabricNames As List(Of String) = DbConnectionManager.GetAllFabricTypeNames()
        Dim colorNames As List(Of String) = DbConnectionManager.GetAllColorNames()

        Dim attributes = New List(Of Object) From {
            New With {.name = "Fabric", .variation = True, .visible = True, .options = fabricNames.ToArray()},
            New With {.name = "Color", .variation = True, .visible = True, .options = colorNames.ToArray()}
        }

        Dim productPayload = New With {
            .name = $"{manufacturerName} {seriesName} {modelName}",
            .type = "variable",
            .sku = parentSku,
            .categories = If(wooCategoryId > 0, New Object() {New With {.id = wooCategoryId}}, New Object() {}),
            .attributes = attributes
        }

        ' 4) Create/Upsert parent product
        Dim productReq As String = JsonSerializer.Serialize(productPayload, JsonOptions)
        Dim productHttp = Await WooCommerceAPI.CreateProductAsync(productPayload)

        Dim wooProductId As Integer
        If productHttp.Status >= 200 AndAlso productHttp.Status < 300 Then
            wooProductId = ExtractIdFromJson(productHttp.Body)
            DbConnectionManager.InsertMpWooSyncLog("upsert_product_parent", parentSku, wooProductId, True, productHttp.Status.ToString(), productReq, productHttp.Body, Nothing)
        Else
            DbConnectionManager.InsertMpWooSyncLog("upsert_product_parent", parentSku, Nothing, False, productHttp.Status.ToString(), productReq, productHttp.Body, "Create product failed")
            res.Success = False
            res.Message = $"Create product failed: HTTP {productHttp.Status}"
            Return res
        End If

        ' Mirror to your mapping table
        Dim mpWooProductId = DbConnectionManager.UpsertMpWooProduct(modelId, parentSku, wooProductId, wooCategoryId, "OK", "Product upserted")

        ' 5) Variations (Fabric x Color)
        For Each f In fabricNames
            For Each c In colorNames
                Dim childSku = $"{parentSku}-{Sanitize(f)}-{Sanitize(c)}"

                Dim variationPayload = New With {
                    .sku = childSku,
                    .attributes = New Object() {
                        New With {.name = "Fabric", .option = f},
                        New With {.name = "Color", .option = c}
                    }
                }

                Dim vReq = JsonSerializer.Serialize(variationPayload, JsonOptions)
                Dim vHttp = Await WooCommerceAPI.CreateVariationAsync(wooProductId, variationPayload)

                If vHttp.Status >= 200 AndAlso vHttp.Status < 300 Then
                    Dim vId = ExtractIdFromJson(vHttp.Body)
                    DbConnectionManager.InsertMpWooSyncLog("upsert_variation_child", childSku, vId, True, vHttp.Status.ToString(), vReq, vHttp.Body, Nothing)
                    DbConnectionManager.UpsertMpWooVariation(mpWooProductId, childSku, vId, f, c, "OK", "Variation upserted")

                    ' OPTIONAL: if you implemented media helpers; else comment these two lines
                    'Dim imgRef = DbConnectionManager.GetVariationImageRef(wooCategoryId, f, c)
                    'DbConnectionManager.UpsertMpWooVariationMedia(mpWooProductId, childSku, imgRef.Item1, imgRef.Item2)

                    res.VariationResults.Add(New VariationOutcome With {.Sku = childSku, .VariationId = vId})
                Else
                    DbConnectionManager.InsertMpWooSyncLog("upsert_variation_child", childSku, Nothing, False, vHttp.Status.ToString(), vReq, vHttp.Body, "Create variation failed")
                End If
            Next
        Next

        res.Success = True
        res.ProductId = wooProductId
        res.ParentSku = parentSku
        res.Message = $"Product {wooProductId} with {res.VariationResults.Count} variations published."
        Return res
    End Function

End Module


