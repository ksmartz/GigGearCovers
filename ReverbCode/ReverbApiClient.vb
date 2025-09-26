' Purpose: Handles authenticated requests to the Reverb API using a personal access token.
' Dependencies: Imports System.Net.Http, Imports Newtonsoft.Json, NuGet: Newtonsoft.Json
' Current date: 2025-09-25

Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Threading.Tasks
Imports Newtonsoft.Json

Public Class ReverbApiClient
    Private Const BASE_URL As String = "https://api.reverb.com/api/"
    Private ReadOnly accessToken As String

    Public Sub New(token As String)
        accessToken = token
    End Sub

    ' Example: Get current user info from Reverb
    Public Async Function GetUserInfoAsync() As Task(Of String)
        Try
            Using client As New HttpClient()
                client.BaseAddress = New Uri(BASE_URL)
                client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", accessToken)
                client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
                client.DefaultRequestHeaders.Add("Accept-Version", "3.0") ' >>> changed

                Dim response = Await client.GetAsync("my/account")
                Dim responseBody = Await response.Content.ReadAsStringAsync()
                If Not response.IsSuccessStatusCode Then
                    Throw New ApplicationException(
                    $"Error contacting Reverb API: {CInt(response.StatusCode)} ({response.ReasonPhrase}){Environment.NewLine}Response body: {responseBody}"
                )
                End If
                Return responseBody
            End Using
        Catch ex As Exception
            Throw New ApplicationException("Error contacting Reverb API: " & ex.Message, ex)
        End Try
    End Function

    ' Purpose: Uploads a new listing to Reverb with all required and custom fields.
    ' Dependencies: Imports System.Net.Http, Imports Newtonsoft.Json, ReverbListing
    ' Current date: 2025-09-25
    Public Async Function UploadListingAsync(listing As ReverbListing) As Task(Of String)
        ' >>> changed
        Try
            Using client As New HttpClient()
                client.BaseAddress = New Uri(BASE_URL)
                client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", accessToken)
                client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
                client.DefaultRequestHeaders.Add("Accept-Version", "3.0")

                ' Build the payload as required by Reverb API
                Dim payload = New With {
                    .listing = listing
                }
                Dim json = JsonConvert.SerializeObject(payload)
                Using content As New StringContent(json, System.Text.Encoding.UTF8, "application/json")
                    Dim response = Await client.PostAsync("listings", content)
                    Dim responseBody = Await response.Content.ReadAsStringAsync()
                    If Not response.IsSuccessStatusCode Then
                        Throw New ApplicationException(
                            $"Error uploading listing: {CInt(response.StatusCode)} ({response.ReasonPhrase}){Environment.NewLine}Response body: {responseBody}"
                        )
                    End If
                    Return responseBody
                End Using
            End Using
        Catch ex As Exception
            Throw New ApplicationException("Error uploading listing: " & ex.Message, ex)
        End Try
        ' <<< end changed
    End Function
End Class
