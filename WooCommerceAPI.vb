





Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Web.Script.Serialization

Public Module WooCommerceAPI
    Private ReadOnly Serializer As New JavaScriptSerializer()

    ' App.config:
    ' <appSettings>
    '   <add key="WooBaseUrl" value="https://yourstore.com/wp-json/wc/v3" />
    '   <add key="WooKey" value="ck_xxx" />
    '   <add key="WooSecret" value="cs_xxx" />
    ' </appSettings>
    Private ReadOnly BaseUrl As String = Configuration.ConfigurationManager.AppSettings("https://giggearcovers.com/wp-json/wc/v3/")
    Private ReadOnly Key As String = Configuration.ConfigurationManager.AppSettings("ck_00941a4c195d8295f746df4e8482bf4c7b4e6c2a")
    Private ReadOnly Secret As String = Configuration.ConfigurationManager.AppSettings("cs_050a34e48ddedae142cc08bb7d4cb24443e44a7a")


    Private Function NewClient() As HttpClient
        Dim c = New HttpClient()
        Dim auth = Convert.ToBase64String(Encoding.UTF8.GetBytes(Key & ":" & Secret))
        c.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Basic", auth)
        c.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
        Return c
    End Function

    Private Function ToJson(o As Object) As String
        Return Serializer.Serialize(o)
    End Function

    Private Async Function ParseJsonAsync(resp As HttpResponseMessage) As Task(Of String)
        Return Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
    End Function

    ' -------- Products --------
    Public Async Function CreateProductAsync(payload As Object) As Task(Of (Status As Integer, Body As String))
        Using c = NewClient()
            Dim json = ToJson(payload)
            Dim resp = Await c.PostAsync($"{BaseUrl}/products", New StringContent(json, Encoding.UTF8, "application/json"))
            Dim body = Await ParseJsonAsync(resp)
            Return (CInt(resp.StatusCode), body)
        End Using
    End Function

    Public Async Function UpdateProductAsync(productId As Integer, payload As Object) As Task(Of (Status As Integer, Body As String))
        Using c = NewClient()
            Dim json = ToJson(payload)
            Dim resp = Await c.PutAsync($"{BaseUrl}/products/{productId}", New StringContent(json, Encoding.UTF8, "application/json"))
            Dim body = Await ParseJsonAsync(resp)
            Return (CInt(resp.StatusCode), body)
        End Using
    End Function

    ' -------- Variations --------
    Public Async Function CreateVariationAsync(productId As Integer, payload As Object) As Task(Of (Status As Integer, Body As String))
        Using c = NewClient()
            Dim json = ToJson(payload)
            Dim resp = Await c.PostAsync($"{BaseUrl}/products/{productId}/variations", New StringContent(json, Encoding.UTF8, "application/json"))
            Dim body = Await ParseJsonAsync(resp)
            Return (CInt(resp.StatusCode), body)
        End Using
    End Function

    Public Async Function UpdateVariationAsync(productId As Integer, variationId As Integer, payload As Object) As Task(Of (Status As Integer, Body As String))
        Using c = NewClient()
            Dim json = ToJson(payload)
            Dim resp = Await c.PutAsync($"{BaseUrl}/products/{productId}/variations/{variationId}", New StringContent(json, Encoding.UTF8, "application/json"))
            Dim body = Await ParseJsonAsync(resp)
            Return (CInt(resp.StatusCode), body)
        End Using
    End Function
End Module

' Defaults (override via App.config <appSettings>)
'Private ReadOnly DefaultWooBaseUrl As String = "https://giggearcovers.com/wp-json/wc/v3/"
'Private ReadOnly DefaultConsumerKey As String = "ck_00941a4c195d8295f746df4e8482bf4c7b4e6c2a"
'Private ReadOnly DefaultConsumerSecret As String = "cs_050a34e48ddedae142cc08bb7d4cb24443e44a7a"
