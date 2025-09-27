' Purpose: Shows a preview of a Reverb listing in txtReverbPreview and returns DialogResult.
' Dependencies: Imports Newtonsoft.Json, ReverbListing, System.Windows.Forms
' Current date: 2025-09-26

Imports Newtonsoft.Json
Imports System.Windows.Forms
Imports System.Drawing
Imports ReverbCode.ReverbListingModel

Public Class MpListingsPreview
    Inherits Form

    ' Holds the listing passed to the preview form
    Private _listing As ReverbListing
    Private lblInfo As Label
    Private flowImages As FlowLayoutPanel
    Private txtDescription As TextBox

    ' Purpose: Constructor that accepts a ReverbListing and sets up the preview UI, including a scrollable description field.
    ' Dependencies: Imports ReverbCode.ReverbListingModel, System.Windows.Forms, System.Drawing
    ' Current date: 2025-09-26

    Public Sub New(listing As ReverbListing)
        _listing = listing
        Me.Text = "Listing Preview"
        Me.Size = New Size(1280, 1024) ' >>> changed

        ' Info label for text fields (excluding description)
        lblInfo = New Label() With {
        .AutoSize = False,
        .Width = 760,
        .Height = 250,
        .Location = New Point(20, 20),
        .Font = New Font("Segoe UI", 10),
        .BorderStyle = BorderStyle.FixedSingle
    }
        Me.Controls.Add(lblInfo)

        ' Scrollable, multi-line TextBox for Description
        txtDescription = New TextBox() With {
        .Multiline = True,
        .ScrollBars = ScrollBars.Vertical,
        .Width = 1024,
        .Height = 500,
        .Location = New Point(20, 275),
        .Font = New Font("Segoe UI", 10),
        .ReadOnly = True
    }
        Me.Controls.Add(txtDescription)

        ' FlowLayoutPanel for images
        flowImages = New FlowLayoutPanel() With {
        .Location = New Point(20, 800),
        .Size = New Size(760, 300),
        .AutoScroll = True
    }
        Me.Controls.Add(flowImages)

        DisplayListingInfo()
        DisplayListingImages()
        DisplayDescription()
    End Sub

    ' Purpose: Displays all listing fields except description in the preview label.
    ' Dependencies: Imports System.Windows.Forms, System.Text
    ' Current date: 2025-09-26
    Private Sub DisplayListingInfo()
        ' >>> changed
        Dim info As New System.Text.StringBuilder()
        info.AppendLine($"Title: {_listing.Title}")
        info.AppendLine($"Condition: {_listing.Condition}")
        info.AppendLine($"Inventory: {_listing.Inventory}")
        info.AppendLine($"SKU: {_listing.Sku}")
        info.AppendLine($"Make: {_listing.Make}")
        info.AppendLine($"Model: {_listing.Model}")
        info.AppendLine($"Price: {_listing.Price:C2}")
        info.AppendLine($"Product Type: {_listing.ProductType}")
        info.AppendLine($"Subcategory 1: {_listing.Subcategory1}")
        info.AppendLine($"Offers Enabled: {_listing.OffersEnabled}")
        info.AppendLine($"Local Pickup: {_listing.LocalPickup}")
        info.AppendLine($"Shipping Profile Name: {_listing.ShippingProfileName}")
        info.AppendLine($"UPC Does Not Apply: {_listing.UpcDoesNotApply}")
        info.AppendLine($"Country Of Origin: {If(String.IsNullOrWhiteSpace(_listing.CountryOfOrigin), "(not set)", _listing.CountryOfOrigin)}") ' >>> changed
        lblInfo.Text = info.ToString()
        ' <<< end changed
    End Sub

    ' Purpose: Displays the description in a scrollable, multi-line TextBox.
    ' Dependencies: Imports System.Windows.Forms
    ' Current date: 2025-09-26
    Private Sub DisplayDescription()
        ' >>> changed
        txtDescription.Text = _listing.Description
        ' <<< end changed
    End Sub

    ' Purpose: Displays all image URLs from the listing.Photos property in the FlowLayoutPanel as labeled clickable links.
    ' Dependencies: Imports System.Windows.Forms, System.Drawing, System.Diagnostics
    ' Current date: 2025-09-26
    Private Sub DisplayListingImages()
        flowImages.Controls.Clear()
        If _listing.Photos IsNot Nothing AndAlso _listing.Photos.Count > 0 Then
            Dim idx As Integer = 1
            For Each url As String In _listing.Photos
                Dim panel As New Panel() With {
                    .Width = 740,
                    .Height = 32,
                    .Margin = New Padding(4)
                }
                Dim label As New Label() With {
                    .Text = $"Image {idx}:",
                    .AutoSize = True,
                    .Font = New Font("Segoe UI", 10, FontStyle.Bold),
                    .Location = New Point(0, 6)
                }
                Dim link As New LinkLabel() With {
                    .Text = url,
                    .AutoSize = True,
                    .LinkColor = Color.Blue,
                    .Font = New Font("Segoe UI", 10),
                    .Location = New Point(label.Width + 8, 6)
                }
                AddHandler link.LinkClicked, Sub(sender As Object, e As LinkLabelLinkClickedEventArgs)
                                                 Try
                                                     Process.Start(New ProcessStartInfo(url) With {.UseShellExecute = True})
                                                 Catch ex As Exception
                                                     MessageBox.Show("Could not open link: " & url, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                 End Try
                                             End Sub
                panel.Controls.Add(label)
                panel.Controls.Add(link)
                flowImages.Controls.Add(panel)
                idx += 1
            Next
        Else
            Dim noImagesLabel As New Label() With {
                .Text = "No images available.",
                .ForeColor = Color.Gray,
                .Width = 200,
                .Height = 30
            }
            flowImages.Controls.Add(noImagesLabel)
        End If
    End Sub

    Private Sub PopulateDescriptionPreview(listing As ReverbListing)
        Try
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
        Catch ex As Exception
            MessageBox.Show("Error displaying description preview: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub MpListingsPreview_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class



