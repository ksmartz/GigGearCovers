' Purpose: Shows a preview of a Reverb listing in txtReverbPreview and returns DialogResult.
' Dependencies: Imports Newtonsoft.Json, ReverbListing, System.Windows.Forms
' Current date: 2025-09-25

Imports Newtonsoft.Json
Imports System.Windows.Forms ' >>> required for IWin32Window and DialogResult

Public Class MpListingsPreview
    ' >>> changed
    ''' <summary>
    ''' Shows the preview form for a given ReverbListing.
    ''' </summary>
    ''' <param name="listing">The ReverbListing to preview.</param>
    ''' <returns>DialogResult.OK if user confirms, DialogResult.Cancel otherwise.</returns>
    Public Function ShowReverbListingPreview(listing As ReverbListing, owner As IWin32Window) As DialogResult
        ' Format the listing as indented JSON for clarity
        txtReverbPreview.Text = JsonConvert.SerializeObject(listing, Formatting.Indented)
        Return Me.ShowDialog(owner)
    End Function

    Public Sub New(listing As ReverbListing)
        InitializeComponent()
        ' Show the ReverbListing as indented JSON in the preview textbox
        If listing IsNot Nothing Then
            txtReverbPreview.Text = JsonConvert.SerializeObject(listing, Formatting.Indented)
        Else
            txtReverbPreview.Text = "No listing data available."
        End If
    End Sub

    ' <<< end changed
    Private Sub frmWooSampleProductPreview_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub







    '*****************************************************************************************************************************************************************
#Region "*************START*************DISPLAY PRODUCT LISTING SAMPLES********************************************************************************************"
    '*****************************************************************************************************************************************************************
#Region "Sample Product Listing Display"

    Public Sub SetWooProductInfo(wooTitle As String, wooLongDescription As String)
        lblWooSampleTitle.Text = wooTitle
        txtReverbPreview.Text = wooLongDescription
        wbWooLongDescriptionPreview.DocumentText = wooLongDescription
    End Sub

    Private Sub txtWooLongDescriptionPreview_TextChanged(sender As Object, e As EventArgs) Handles txtReverbPreview.TextChanged

    End Sub

#End Region


#End Region '---------END---------------BUTTON EVENT HANDLERS---------------------------------------------------------------------------------------------------------
End Class