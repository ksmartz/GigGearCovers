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

    ' =====================================================================
    ' Sub: UpdateWooImageIds
    ' Date: 2025-09-06
    ' Purpose:
    '   Persist WooCommerce image IDs for a given model/equipment type.
    '   Fixes the SQL error "Cannot insert NULL into column 'FK_equipmentTypeId'..."
    '   by resolving a valid equipmentTypeId before insert/update.
    ' Parameters:
    '   conn             : open SqlConnection
    '   tx               : active SqlTransaction (optional but recommended)
    '   modelId          : your internal model ID
    '   productResponseJson : raw JSON from Woo product create/update/get
    '   productTypeName  : fallback lookup key for EquipmentType.Name if Model.FK_equipmentTypeId is NULL
    ' =====================================================================
    ' =============================================================================================
    ' OVERLOAD 2 (with transaction): original signature
    ' =============================================================================================
    Public  Sub UpdateWooImageIds(conn As SqlConnection,
                                        tx As SqlTransaction,
                                        modelId As Integer,
                                        productResponseJson As String,
                                        productTypeName As String)

        ' Parse the image id from Woo response
        Dim wooImageId As Long = ParsePrimaryWooImageId(productResponseJson)

        ' Resolve equipment type id (prevents NULL FK_equipmentTypeId errors)
        Dim equipmentTypeId As Integer? = ResolveEquipmentTypeId(conn, tx, modelId, productTypeName)
        If Not equipmentTypeId.HasValue Then
            Throw New InvalidOperationException(
                $"Cannot save Woo image because EquipmentTypeId is unknown. modelId={modelId}, productType='{productTypeName}'.")
        End If

        ' UPSERT row in ModelEquipmentTypeImage
        Using cmd As SqlCommand = NewCmd("
            IF EXISTS (SELECT 1 FROM dbo.ModelEquipmentTypeImage WITH (UPDLOCK, HOLDLOCK)
                       WHERE FK_modelId=@modelId AND FK_equipmentTypeId=@equipmentTypeId)
            BEGIN
                UPDATE dbo.ModelEquipmentTypeImage
                   SET WooImageId = @wooImageId,
                       UpdatedAt  = SYSUTCDATETIME()
                 WHERE FK_modelId=@modelId AND FK_equipmentTypeId=@equipmentTypeId;
            END
            ELSE
            BEGIN
                INSERT INTO dbo.ModelEquipmentTypeImage
                    (FK_modelId, FK_equipmentTypeId, WooImageId, CreatedAt)
                VALUES
                    (@modelId, @equipmentTypeId, @wooImageId, SYSUTCDATETIME());
            END
        ", conn, tx)
            cmd.Parameters.Add("@modelId", SqlDbType.Int).Value = modelId
            cmd.Parameters.Add("@equipmentTypeId", SqlDbType.Int).Value = equipmentTypeId.Value
            cmd.Parameters.Add("@wooImageId", SqlDbType.BigInt).Value = wooImageId
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    ' =====================================================================
    ' Function: ParsePrimaryWooImageId
    ' Date: 2025-09-06
    ' Purpose:
    '   Extract the primary image ID from a WooCommerce product JSON payload.
    '   Looks for product.images[0].id. Returns 0 if not found.
    ' =====================================================================
    Public Function ParsePrimaryWooImageId(json As String) As Long
        If String.IsNullOrWhiteSpace(json) Then Return 0

        Dim token As JToken = Nothing
        Try
            token = JToken.Parse(json)
        Catch
            Return 0
        End Try

        Dim images As JToken = Nothing

        If token.Type = JTokenType.Object Then
            images = token("images")
        ElseIf token.Type = JTokenType.Array Then
            ' If an array of products is returned, check first element
            Dim first = token.First
            If first IsNot Nothing Then images = first("images")
        End If

        If images IsNot Nothing AndAlso images.Type = JTokenType.Array AndAlso images.HasValues Then
            Dim firstImg = images.First
            If firstImg IsNot Nothing Then
                Dim idTok = firstImg("id")
                If idTok IsNot Nothing AndAlso idTok.Type <> JTokenType.Null Then
                    Dim val As Long
                    If Long.TryParse(idTok.ToString(), val) Then Return val
                End If
            End If
        End If

        Return 0
    End Function

    ' =====================================================================
    ' Function: ParseWooProductIdFromResult
    ' Date: 2025-09-06
    ' Purpose:
    '   Extract the WooCommerce product "id" from a JSON response payload.
    '   Works for typical Woo endpoints (create/update/get single product).
    ' Returns:
    '   Product ID as Long. Returns 0 if not found or if JSON is empty.
    ' =====================================================================
    Public Function ParseWooProductIdFromResult(json As String) As Long
        If String.IsNullOrWhiteSpace(json) Then Return 0

        Dim token As JToken = Nothing
        Try
            token = JToken.Parse(json)
        Catch
            ' Not valid JSON
            Return 0
        End Try

        ' Response can be an object with "id" or an array (rare). Prefer object->id.
        If token.Type = JTokenType.Object Then
            Dim idTok = token("id")
            If idTok IsNot Nothing AndAlso idTok.Type <> JTokenType.Null Then
                Dim val As Long
                If Long.TryParse(idTok.ToString(), val) Then Return val
            End If
        ElseIf token.Type = JTokenType.Array Then
            ' e.g., some bulk endpoints might return an array of products
            Dim first = token.First
            If first IsNot Nothing Then
                Dim idTok = first("id")
                If idTok IsNot Nothing AndAlso idTok.Type <> JTokenType.Null Then
                    Dim val As Long
                    If Long.TryParse(idTok.ToString(), val) Then Return val
                End If
            End If
        End If

        Return 0
    End Function
    ' =====================================================================
    ' Helper: ResolveEquipmentTypeId
    ' Date: 2025-09-06
    ' Purpose:
    '   Determine EquipmentTypeId for a model.
    '   1) Try Model.FK_equipmentTypeId
    '   2) Try lookup by EquipmentType.Name = productTypeName
    ' Returns:
    '   Integer? (Nothing if not found)
    ' =====================================================================
    Private Function ResolveEquipmentTypeId(
        ByVal conn As SqlConnection,
        ByVal tx As SqlTransaction,
        ByVal modelId As Integer,
        ByVal productTypeName As String
    ) As Integer?

        ' 1) From Model row
        Using cmd As New SqlCommand("
            SELECT TOP(1) FK_equipmentTypeId
            FROM dbo.Model
            WHERE PK_modelId = @modelId
        ", conn, tx)
            cmd.Parameters.Add("@modelId", SqlDbType.Int).Value = modelId
            Dim obj = cmd.ExecuteScalar()
            If obj IsNot Nothing AndAlso obj IsNot DBNull.Value Then
                Dim id As Integer = Convert.ToInt32(obj)
                If id > 0 Then Return id
            End If
        End Using

        ' 2) From EquipmentType.Name
        If Not String.IsNullOrWhiteSpace(productTypeName) Then
            Using cmd As New SqlCommand("
                SELECT TOP(1) PK_equipmentTypeId
                FROM dbo.EquipmentType
                WHERE Name = @name
            ", conn, tx)
                cmd.Parameters.Add("@name", SqlDbType.NVarChar, 100).Value = productTypeName.Trim()
                Dim obj = cmd.ExecuteScalar()
                If obj IsNot Nothing AndAlso obj IsNot DBNull.Value Then
                    Dim id As Integer = Convert.ToInt32(obj)
                    If id > 0 Then Return id
                End If
            End Using
        End If

        Return Nothing
    End Function


End Module

