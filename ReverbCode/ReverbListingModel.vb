' Purpose: Represents a Reverb listing payload for upload.
' Dependencies: Imports System.Collections.Generic, Newtonsoft.Json
' Current date: 2025-09-25
' Purpose: Adds a FieldNameValues property to return all field names and their values as a dictionary.
' Dependencies: Imports System.Collections.Generic, Newtonsoft.Json
' Current date: 2025-09-26

Imports Newtonsoft.Json
Imports System.Collections.Generic

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
    Public Property Photos As List(Of String)

    ' >>> changed
    ' Returns a dictionary of all property names and their values.
    Public ReadOnly Property FieldNameValues As Dictionary(Of String, Object)
        Get
            Dim dict As New Dictionary(Of String, Object)()
            dict("Title") = Title
            dict("Condition") = Condition
            dict("Inventory") = Inventory
            dict("Sku") = Sku
            dict("Make") = Make
            dict("Model") = Model
            dict("Description") = Description
            dict("Price") = Price
            dict("ProductType") = ProductType
            dict("Subcategory1") = Subcategory1
            dict("OffersEnabled") = OffersEnabled
            dict("LocalPickup") = LocalPickup
            dict("ShippingProfileName") = ShippingProfileName
            dict("UpcDoesNotApply") = UpcDoesNotApply
            dict("CountryOfOrigin") = CountryOfOrigin
            dict("Photos") = Photos
            Return dict
        End Get
    End Property
    ' <<< end changed

End Class


