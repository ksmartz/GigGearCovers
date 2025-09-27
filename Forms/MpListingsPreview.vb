' Purpose: Shows a preview of a Reverb listing in txtReverbPreview and returns DialogResult.
' Dependencies: Imports Newtonsoft.Json, ReverbListing, System.Windows.Forms
' Current date: 2025-09-26

Imports Newtonsoft.Json
Imports System.Windows.Forms
Imports ReverbCode.ReverbListingModel

Public Class MpListingsPreview
    Inherits Form

    ' >>> changed
    ' Holds the listing passed to the preview form
    Private _listing As ReverbListing
    ' <<< end changed

    ' Purpose: Constructor that accepts a ReverbListing and stores it for use in the form.
    ' Dependencies: Imports ReverbCode.ReverbListingModel
    ' Current date: 2025-09-26
    Public Sub New(listing As ReverbListing)
        InitializeComponent()
        ' >>> changed
        _listing = listing
        ' <<< end changed
        PopulateDescriptionPreview(_listing)
    End Sub

    Private Sub PopulateDescriptionPreview(listing As ReverbListing)
        Try
            ' >>> changed
            ' Build a full preview of all listing fields
            Dim previewText As New System.Text.StringBuilder()
            previewText.AppendLine($"Title: {listing.Title}")
            previewText.AppendLine($"Condition: {listing.Condition}")
            previewText.AppendLine($"Inventory: {listing.Inventory}")
            previewText.AppendLine($"SKU: {listing.Sku}")
            previewText.AppendLine($"Make: {listing.Make}")
            previewText.AppendLine($"Model: {listing.Model}")
            previewText.AppendLine($"Price: {listing.Price:C2}")
            previewText.AppendLine($"Product Type: {listing.ProductType}")
            previewText.AppendLine($"Subcategory 1: {listing.Subcategory1}")
            previewText.AppendLine($"Offers Enabled: {listing.OffersEnabled}")
            previewText.AppendLine($"Local Pickup: {listing.LocalPickup}")
            previewText.AppendLine($"Shipping Profile Name: {listing.ShippingProfileName}")
            previewText.AppendLine($"UPC Does Not Apply: {listing.UpcDoesNotApply}")
            previewText.AppendLine($"Country Of Origin: {listing.CountryOfOrigin}")
            previewText.AppendLine("Description:")
            previewText.AppendLine(listing.Description)
            txtDescriptionPreview.Text = previewText.ToString()
            ' <<< end changed
        Catch ex As Exception
            MessageBox.Show("Error displaying description preview: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class



