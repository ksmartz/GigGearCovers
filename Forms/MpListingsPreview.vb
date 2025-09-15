Public Class MpListingsPreview
    Private Sub frmWooSampleProductPreview_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub







    '*****************************************************************************************************************************************************************
#Region "*************START*************DISPLAY PRODUCT LISTING SAMPLES********************************************************************************************"
    '*****************************************************************************************************************************************************************
#Region "Sample Product Listing Display"

    Public Sub SetWooProductInfo(wooTitle As String, wooLongDescription As String)
        lblWooSampleTitle.Text = wooTitle
        txtWooLongDescriptionPreview.Text = wooLongDescription
        wbWooLongDescriptionPreview.DocumentText = wooLongDescription
    End Sub

    Private Sub txtWooLongDescriptionPreview_TextChanged(sender As Object, e As EventArgs) Handles txtWooLongDescriptionPreview.TextChanged

    End Sub

#End Region


#End Region '---------END---------------BUTTON EVENT HANDLERS---------------------------------------------------------------------------------------------------------
End Class