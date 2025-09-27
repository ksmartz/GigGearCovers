' Purpose: Handles authenticated requests to the Reverb API using a personal access token.
' Dependencies: Imports System.Net.Http, Imports Newtonsoft.Json, NuGet: Newtonsoft.Json
' Current date: 2025-09-25

Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json

Public Class ReverbApiClient
    Private Const BASE_URL As String = "https://api.reverb.com/api/"
    Private ReadOnly accessToken As String

    Public Sub New(token As String)
        accessToken = token
    End Sub
    ' Purpose: Checks if a SKU exists on Reverb by querying the Reverb API, and displays a MessageBox indicating the result.
    ' Dependencies: Imports System.Net.Http, Imports System.Net.Http.Headers, Imports Newtonsoft.Json, Imports System.Windows.Forms
    ' Current date: 2025-09-27
    Public Async Function SkuExistsOnReverbAsync(sku As String) As Task(Of Boolean)
        ' >>> changed
        ' Build the Reverb API search URL for the SKU, including all states (live + draft)
        Dim url As String = $"{BASE_URL}my/listings?sku={Uri.EscapeDataString(sku)}&state=all"

        Try
            Using client As New HttpClient()
                ' Set required headers for Reverb API
                client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", accessToken)
                client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
                client.DefaultRequestHeaders.Add("Accept-Version", "3.0") ' Required header

                ' Send GET request to Reverb API
                Dim response As HttpResponseMessage = Await client.GetAsync(url)
                Dim json As String = Await response.Content.ReadAsStringAsync()
                Debug.WriteLine("Reverb API raw response: " & json)

                If Not response.IsSuccessStatusCode Then
                    Throw New Exception($"Reverb API error: {response.StatusCode} - {response.ReasonPhrase}{vbCrLf}Response: {json}")
                End If

                ' Deserialize the JSON response to check for listings
                Dim result = JsonConvert.DeserializeObject(Of ReverbListingsResponse)(json)
                If result IsNot Nothing AndAlso result.listings IsNot Nothing AndAlso result.listings.Count > 0 Then
                    MessageBox.Show($"SKU '{sku}' was found on Reverb.", "SKU Check Result", MessageBoxButtons.OK, MessageBoxIcon.Information) ' >>> changed
                    Return True ' SKU found
                Else
                    MessageBox.Show($"SKU '{sku}' was NOT found on Reverb.", "SKU Check Result", MessageBoxButtons.OK, MessageBoxIcon.Information) ' >>> changed
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("Error checking SKU on Reverb: " & ex.Message, "Reverb API Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine("Exception: " & ex.ToString())
        End Try

        Return False ' SKU not found or error
        ' <<< end changed
    End Function

    ' Model for deserializing Reverb API response
    Private Class ReverbListingsResponse
        <JsonProperty("listings")>
        Public Property listings As List(Of Object)
    End Class
    ' Purpose: Uploads a listing to Reverb using the required JSON format.
    ' Dependencies: Imports System.Net.Http, Imports System.Net.Http.Headers, Imports Newtonsoft.Json, Imports System.Windows.Forms
    ' Current date: 2025-09-27
    Public Async Function UploadListingToReverbAsync(listing As ReverbListing) As Task(Of Boolean)
        ' >>> changed
        ' Build the Reverb API endpoint for creating a listing
        Dim url As String = $"{BASE_URL}my/listings"

        Try
            Using client As New HttpClient()
                ' Set required headers for Reverb API
                client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", accessToken)
                client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
                client.DefaultRequestHeaders.Add("Accept-Version", "3.0") ' Required header

                ' Build the JSON payload in the format Reverb expects
                Dim payload As New Dictionary(Of String, Object) From {
                    {"listing", listing}
                }
                Dim jsonPayload As String = JsonConvert.SerializeObject(payload)

                ' Send POST request to Reverb API
                Dim content As New StringContent(jsonPayload, Text.Encoding.UTF8, "application/json")
                Dim response As HttpResponseMessage = Await client.PostAsync(url, content)
                Dim jsonResponse As String = Await response.Content.ReadAsStringAsync()
                Debug.WriteLine("Reverb API upload response: " & jsonResponse)

                If response.IsSuccessStatusCode Then
                    MessageBox.Show("Listing uploaded to Reverb successfully!", "Upload Result", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return True
                Else
                    MessageBox.Show($"Failed to upload listing to Reverb.{vbCrLf}Status: {response.StatusCode} {response.ReasonPhrase}{vbCrLf}Response: {jsonResponse}", "Upload Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("Error uploading listing to Reverb: " & ex.Message, "Upload Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine("Exception: " & ex.ToString())
        End Try

        Return False
        ' <<< end changed
    End Function

End Class
