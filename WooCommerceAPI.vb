Imports System.Data.SqlClient
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Module WooCommerceApi
    Private ReadOnly WooCommerceUrl As String = "https://giggearcovers.com/wp-json/wc/v3/"
    Private ReadOnly ConsumerKey As String = "ck_00941a4c195d8295f746df4e8482bf4c7b4e6c2a"
    Private ReadOnly ConsumerSecret As String = "cs_050a34e48ddedae142cc08bb7d4cb24443e44a7a"

    ' Helper to create an authenticated HttpClient
    Private Function CreateClient() As HttpClient
        Dim client As New HttpClient()
        client.BaseAddress = New Uri(WooCommerceUrl)
        Dim auth = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{ConsumerKey}:{ConsumerSecret}"))
        client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Basic", auth)
        Return client
    End Function

    ' Serializes the product object, omitting nulls (such as id for new images)
    Public Function SerializeProduct(product As MpWooCommerceProduct) As String
        Dim settings As New JsonSerializerSettings With {
            .NullValueHandling = NullValueHandling.Ignore
        }
        Return JsonConvert.SerializeObject(product, settings)
    End Function

    ' Uploads a product (parent or variation) - ensures correct image serialization
    Public Async Function UploadProductAsync(product As MpWooCommerceProduct, Optional endpoint As String = "products") As Task(Of String)
        Dim jsonBody As String = SerializeProduct(product)
        Using client = CreateClient()
            Dim content = New StringContent(jsonBody, Encoding.UTF8, "application/json")
            Dim response = Await client.PostAsync(endpoint, content)
            Dim result = Await response.Content.ReadAsStringAsync()
            If Not response.IsSuccessStatusCode Then
                Throw New Exception($"API Error: {response.StatusCode} - {result}")
            End If
            Return result
        End Using
    End Function

    ' Updates a product by WooCommerce product ID (PUT) - ensures correct image serialization
    Public Async Function UpdateProductAsync(product As MpWooCommerceProduct, wooProductId As Integer) As Task(Of String)
        Dim jsonBody As String = SerializeProduct(product)
        Using client = CreateClient()
            Dim content = New StringContent(jsonBody, Encoding.UTF8, "application/json")
            Dim response = Await client.PutAsync($"products/{wooProductId}", content)
            Dim result = Await response.Content.ReadAsStringAsync()
            If Not response.IsSuccessStatusCode Then
                Throw New Exception($"API Error: {response.StatusCode} - {result}")
            End If
            Return result
        End Using
    End Function

    ' Uploads a product from a JSON string (deserializes to MpWooCommerceProduct)
    Public Async Function UploadProductFromJsonAsync(jsonBody As String, Optional endpoint As String = "products") As Task(Of String)
        Dim product As MpWooCommerceProduct = JsonConvert.DeserializeObject(Of MpWooCommerceProduct)(jsonBody)
        Return Await UploadProductAsync(product, endpoint)
    End Function

    ' Updates a product from a JSON string (deserializes to MpWooCommerceProduct)
    Public Async Function UpdateProductFromJsonAsync(jsonBody As String, wooProductId As Integer) As Task(Of String)
        Dim product As MpWooCommerceProduct = JsonConvert.DeserializeObject(Of MpWooCommerceProduct)(jsonBody)
        Return Await UpdateProductAsync(product, wooProductId)
    End Function

    ' Example function to update image IDs in the database for marketplace ID 6
    Public Sub UpdateWooImageIds(productResponseJson As String)
        Dim productResponse = JObject.Parse(productResponseJson)
        Dim images = productResponse("images")
        If images Is Nothing OrElse Not TypeOf images Is JArray Then Return

        Using conn As SqlConnection = DbConnectionManager.GetConnection()
            conn.Open()
            For Each img In images
                Dim wooImageId As Integer = img("id")
                Dim imageUrl As String = img("src").ToString()

                ' Update or insert the image record for marketplace ID 6
                Dim cmd As New SqlCommand("
                UPDATE ProductImages
                SET WooImageId = @WooImageId
                WHERE MarketplaceId = 6 AND ImageUrl = @ImageUrl

                IF @@ROWCOUNT = 0
                INSERT INTO ProductImages (MarketplaceId, WooImageId, ImageUrl)
                VALUES (6, @WooImageId, @ImageUrl)
            ", conn)
                cmd.Parameters.AddWithValue("@WooImageId", wooImageId)
                cmd.Parameters.AddWithValue("@ImageUrl", imageUrl)
                cmd.ExecuteNonQuery()
            Next
        End Using
    End Sub

    ' Parses the WooCommerce product ID from the API response JSON
    Public Function ParseWooProductIdFromResult(resultJson As String) As Integer
        Dim obj = JObject.Parse(resultJson)
        If obj("id") IsNot Nothing Then
            Return CInt(obj("id"))
        End If
        Throw New Exception("Could not parse WooCommerce product ID from response.")
    End Function
End Module

