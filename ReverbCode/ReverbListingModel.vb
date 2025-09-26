' Purpose: Represents a Reverb listing payload for upload.
' Dependencies: Imports System.Collections.Generic, Newtonsoft.Json
' Current date: 2025-09-25

Imports Newtonsoft.Json

Public Class ReverbListing
    <JsonProperty("title")>
    Public Property Title As String

    <JsonProperty("condition")>
    Public Property Condition As String

    <JsonProperty("inventory")>
    Public Property Inventory As Integer

    <JsonProperty("sku")>
    Public Property Sku As String

    <JsonProperty("make")>
    Public Property Make As String

    <JsonProperty("model")>
    Public Property Model As String

    <JsonProperty("description")>
    Public Property Description As String

    <JsonProperty("price")>
    Public Property Price As Decimal

    <JsonProperty("product_type")>
    Public Property ProductType As String

    <JsonProperty("subcategory_1")>
    Public Property Subcategory1 As String

    <JsonProperty("offers_enabled")>
    Public Property OffersEnabled As Boolean

    <JsonProperty("local_pickup")>
    Public Property LocalPickup As Boolean

    <JsonProperty("shipping_profile_name")>
    Public Property ShippingProfileName As String

    <JsonProperty("upc_does_not_apply")>
    Public Property UpcDoesNotApply As Boolean

    <JsonProperty("country_of_origin")>
    Public Property CountryOfOrigin As String

    <JsonProperty("photos")>
    Public Property Photos As List(Of ReverbPhoto)
End Class

Public Class ReverbPhoto
    <JsonProperty("photo_url")>
    Public Property PhotoUrl As String
End Class
