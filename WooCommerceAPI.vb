





Imports System
Imports System.Configuration
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Web.Script.Serialization
'==============================================================================
' Module : WooCommerceAPI
' Purpose: Safe, absolute-URI WooCommerce v3 client with tuple results.
'
' App.config (example):
'   <appSettings>
'     <!-- Use the site root OR the wc/v3 root. Either is fine: -->
'     <!-- Option A: site root (module appends wp-json/wc/v3/...) -->
'     <!-- <add key="WooBaseUrl" value="https://yourstore.com/" /> -->
'     <!-- Option B: wc/v3 root -->
'     <add key="WooBaseUrl" value="https://yourstore.com/wp-json/wc/v3/" />
'     <add key="WooKey" value="ck_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx" />
'     <add key="WooSecret" value="cs_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx" />
'   </appSettings>
'
' Notes:
'   - Always builds ABSOLUTE URIs (fixes “invalid request URI”).
'   - Basic Auth (ck:cs) over HTTPS. If your host requires query params instead,
'     we can switch, but this is the most reliable.
'   - Returns (Status As Integer, Body As String).
'
' Date   : 2025-09-15
'------------------------------------------------------------------------------
Public Module WooCommerceAPI

    ' Single shared HttpClient (do not dispose)
    Private ReadOnly _http As New HttpClient()

    ' Config / state
    Private _configured As Boolean = False
    Private _baseApi As Uri = Nothing      ' e.g., https://giggearcovers.com/wp-json/wc/v3/
    Private _ck As String = Nothing
    Private _cs As String = Nothing

    ' JSON serializer (kept to match your project)
    Private ReadOnly _serializer As New JavaScriptSerializer()

    '------------------------------------------------------------------------------
    ' Public API
    '------------------------------------------------------------------------------

    ''' <summary>
    ''' Create (or upsert) a product. Payload can be an anonymous object.
    ''' POST {base}/products
    ''' </summary>
    Public Async Function CreateProductAsync(payload As Object) As Threading.Tasks.Task(Of (Status As Integer, Body As String))
        EnsureConfigured()
        Dim uri As Uri = BuildEndpointUri("products")
        Dim json As String = _serializer.Serialize(payload)
        Using req As New HttpRequestMessage(HttpMethod.Post, uri)
            req.Content = New StringContent(json, Encoding.UTF8, "application/json")
            Using resp As HttpResponseMessage = Await _http.SendAsync(req).ConfigureAwait(False)
                Dim body As String = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
                Return (CInt(resp.StatusCode), body)
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Update an existing product.
    ''' PUT {base}/products/{productId}
    ''' </summary>
    Public Async Function UpdateProductAsync(productId As Integer, payload As Object) As Threading.Tasks.Task(Of (Status As Integer, Body As String))
        If productId <= 0 Then Return (400, "Invalid productId.")
        EnsureConfigured()
        Dim uri As Uri = BuildEndpointUri($"products/{productId}")
        Dim json As String = _serializer.Serialize(payload)
        Using req As New HttpRequestMessage(HttpMethod.Put, uri)
            req.Content = New StringContent(json, Encoding.UTF8, "application/json")
            Using resp As HttpResponseMessage = Await _http.SendAsync(req).ConfigureAwait(False)
                Dim body As String = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
                Return (CInt(resp.StatusCode), body)
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Create (or upsert) a variation for a parent product.
    ''' POST {base}/products/{productId}/variations
    ''' </summary>
    Public Async Function CreateVariationAsync(productId As Integer, payload As Object) As Threading.Tasks.Task(Of (Status As Integer, Body As String))
        If productId <= 0 Then Return (400, "Invalid parent product id.")
        EnsureConfigured()
        Dim uri As Uri = BuildEndpointUri($"products/{productId}/variations")
        Dim json As String = _serializer.Serialize(payload)
        Using req As New HttpRequestMessage(HttpMethod.Post, uri)
            req.Content = New StringContent(json, Encoding.UTF8, "application/json")
            Using resp As HttpResponseMessage = Await _http.SendAsync(req).ConfigureAwait(False)
                Dim body As String = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
                Return (CInt(resp.StatusCode), body)
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Update an existing variation.
    ''' PUT {base}/products/{productId}/variations/{variationId}
    ''' </summary>
    Public Async Function UpdateVariationAsync(productId As Integer, variationId As Integer, payload As Object) As Threading.Tasks.Task(Of (Status As Integer, Body As String))
        If productId <= 0 OrElse variationId <= 0 Then Return (400, "Invalid product/variation id.")
        EnsureConfigured()
        Dim uri As Uri = BuildEndpointUri($"products/{productId}/variations/{variationId}")
        Dim json As String = _serializer.Serialize(payload)
        Using req As New HttpRequestMessage(HttpMethod.Put, uri)
            req.Content = New StringContent(json, Encoding.UTF8, "application/json")
            Using resp As HttpResponseMessage = Await _http.SendAsync(req).ConfigureAwait(False)
                Dim body As String = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
                Return (CInt(resp.StatusCode), body)
            End Using
        End Using
    End Function

    '------------------------------------------------------------------------------
    ' Configuration / Helpers
    '------------------------------------------------------------------------------

    ''' <summary>
    ''' Ensure HttpClient and configuration are initialized. Reads:
    '''   WooBaseUrl, WooKey, WooSecret
    ''' </summary>
    Private Sub EnsureConfigured()
        If _configured Then Return

        Dim rawBase As String = GetAppSetting("WooBaseUrl")
        _ck = GetAppSetting("WooKey")
        _cs = GetAppSetting("WooSecret")

        If String.IsNullOrWhiteSpace(rawBase) Then
            Throw New InvalidOperationException("AppSetting 'WooBaseUrl' is missing.")
        End If
        If String.IsNullOrWhiteSpace(_ck) OrElse String.IsNullOrWhiteSpace(_cs) Then
            Throw New InvalidOperationException("AppSettings 'WooKey' and/or 'WooSecret' are missing.")
        End If

        ' Normalize base. Support either site root or wc/v3 root.
        rawBase = rawBase.Trim()
        ' If user gave the site root, append wc/v3 route.
        If rawBase.EndsWith("/wp-json/wc/v3", StringComparison.OrdinalIgnoreCase) Then rawBase &= "/"
        If rawBase.EndsWith("/wp-json/wc/v3/", StringComparison.OrdinalIgnoreCase) Then
            ' Already at API root; keep as-is
        Else
            ' If they passed the site root, ensure it ends with slash and then append wc/v3
            If Not rawBase.EndsWith("/") Then rawBase &= "/"
            rawBase &= "wp-json/wc/v3/"
        End If

        Dim parsed As Uri = Nothing
        If Not Uri.TryCreate(rawBase, UriKind.Absolute, parsed) Then
            Throw New InvalidOperationException($"WooBaseUrl is not a valid absolute URI: {rawBase}")
        End If
        If parsed.Scheme <> Uri.UriSchemeHttps AndAlso parsed.Scheme <> Uri.UriSchemeHttp Then
            Throw New InvalidOperationException("WooBaseUrl must start with http:// or https://")
        End If
        _baseApi = parsed

        ' Configure HttpClient once
        _http.BaseAddress = _baseApi
        _http.DefaultRequestHeaders.Accept.Clear()
        _http.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))

        ' Basic auth header for ck:cs (preemptive)
        Dim rawCred As String = $"{_ck}:{_cs}"
        Dim b64 As String = Convert.ToBase64String(Encoding.ASCII.GetBytes(rawCred))
        _http.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Basic", b64)

        _configured = True
    End Sub

    Private Function GetAppSetting(key As String) As String
        Dim v As String = ConfigurationManager.AppSettings(key)
        If v Is Nothing Then Return Nothing
        Return v.Trim()
    End Function

    ''' <summary>
    ''' Build an absolute endpoint URI under the configured wc/v3 root.
    ''' e.g., "products/123/variations" → https://.../wp-json/wc/v3/products/123/variations
    ''' </summary>
    Private Function BuildEndpointUri(relativePath As String) As Uri
        If String.IsNullOrWhiteSpace(relativePath) Then
            Throw New ArgumentException("relativePath cannot be empty.")
        End If
        ' Remove any leading slash so New Uri(base, rel) doesn't reset the path.
        Dim rel As String = relativePath.TrimStart("/"c)
        Return New Uri(_baseApi, rel)
    End Function




    ' Add/replace in WooCommerceAPI.vb

    Public Async Function GetProductsBySkuAsync(sku As String) As Threading.Tasks.Task(Of (Status As Integer, Body As String))
        EnsureConfigured()
        Dim endpoint As Uri = BuildEndpointUri("products?sku=" & Uri.EscapeDataString(sku))
        Using req As New HttpRequestMessage(HttpMethod.Get, endpoint)
            Using resp As HttpResponseMessage = Await _http.SendAsync(req).ConfigureAwait(False)
                Dim body As String = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
                Return (CInt(resp.StatusCode), body)
            End Using
        End Using
    End Function

    Public Async Function GetVariationsBySkuAsync(parentProductId As Integer, sku As String) As Threading.Tasks.Task(Of (Status As Integer, Body As String))
        EnsureConfigured()
        Dim rel As String = $"products/{parentProductId}/variations?sku={Uri.EscapeDataString(sku)}"
        Dim endpoint As Uri = BuildEndpointUri(rel)
        Using req As New HttpRequestMessage(HttpMethod.Get, endpoint)
            Using resp As HttpResponseMessage = Await _http.SendAsync(req).ConfigureAwait(False)
                Dim body As String = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
                Return (CInt(resp.StatusCode), body)
            End Using
        End Using
    End Function


End Module


'Public Module WooCommerceAPI
'    Private ReadOnly Serializer As New JavaScriptSerializer()

'    ' App.config:
'    ' <appSettings>
'    '   <add key="WooBaseUrl" value="https://yourstore.com/wp-json/wc/v3" />
'    '   <add key="WooKey" value="ck_xxx" />
'    '   <add key="WooSecret" value="cs_xxx" />
'    ' </appSettings>
'    Private ReadOnly BaseUrl As String = Configuration.ConfigurationManager.AppSettings("https://giggearcovers.com/wp-json/wc/v3/")
'    Private ReadOnly Key As String = Configuration.ConfigurationManager.AppSettings("ck_00941a4c195d8295f746df4e8482bf4c7b4e6c2a")
'    Private ReadOnly Secret As String = Configuration.ConfigurationManager.AppSettings("cs_050a34e48ddedae142cc08bb7d4cb24443e44a7a")


'    Private Function NewClient() As HttpClient
'        Dim c = New HttpClient()
'        Dim auth = Convert.ToBase64String(Encoding.UTF8.GetBytes(Key & ":" & Secret))
'        c.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Basic", auth)
'        c.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
'        Return c
'    End Function

'    Private Function ToJson(o As Object) As String
'        Return Serializer.Serialize(o)
'    End Function

'    Private Async Function ParseJsonAsync(resp As HttpResponseMessage) As Task(Of String)
'        Return Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
'    End Function

'    ' -------- Products --------
'    Public Async Function CreateProductAsync(payload As Object) As Task(Of (Status As Integer, Body As String))
'        Using c = NewClient()
'            Dim json = ToJson(payload)
'            Dim resp = Await c.PostAsync($"{BaseUrl}/products", New StringContent(json, Encoding.UTF8, "application/json"))
'            Dim body = Await ParseJsonAsync(resp)
'            Return (CInt(resp.StatusCode), body)
'        End Using
'    End Function

'    Public Async Function UpdateProductAsync(productId As Integer, payload As Object) As Task(Of (Status As Integer, Body As String))
'        Using c = NewClient()
'            Dim json = ToJson(payload)
'            Dim resp = Await c.PutAsync($"{BaseUrl}/products/{productId}", New StringContent(json, Encoding.UTF8, "application/json"))
'            Dim body = Await ParseJsonAsync(resp)
'            Return (CInt(resp.StatusCode), body)
'        End Using
'    End Function

'    ' -------- Variations --------
'    Public Async Function CreateVariationAsync(productId As Integer, payload As Object) As Task(Of (Status As Integer, Body As String))
'        Using c = NewClient()
'            Dim json = ToJson(payload)
'            Dim resp = Await c.PostAsync($"{BaseUrl}/products/{productId}/variations", New StringContent(json, Encoding.UTF8, "application/json"))
'            Dim body = Await ParseJsonAsync(resp)
'            Return (CInt(resp.StatusCode), body)
'        End Using
'    End Function

'    Public Async Function UpdateVariationAsync(productId As Integer, variationId As Integer, payload As Object) As Task(Of (Status As Integer, Body As String))
'        Using c = NewClient()
'            Dim json = ToJson(payload)
'            Dim resp = Await c.PutAsync($"{BaseUrl}/products/{productId}/variations/{variationId}", New StringContent(json, Encoding.UTF8, "application/json"))
'            Dim body = Await ParseJsonAsync(resp)
'            Return (CInt(resp.StatusCode), body)
'        End Using
'    End Function
'End Module

' Defaults (override via App.config <appSettings>)
'Private ReadOnly DefaultWooBaseUrl As String = "https://giggearcovers.com/wp-json/wc/v3/"
'Private ReadOnly DefaultConsumerKey As String = "ck_00941a4c195d8295f746df4e8482bf4c7b4e6c2a"
'Private ReadOnly DefaultConsumerSecret As String = "cs_050a34e48ddedae142cc08bb7d4cb24443e44a7a"
