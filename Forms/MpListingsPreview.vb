Public Class MpListingsPreview
    Private Sub frmWooSampleProductPreview_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub







    '*****************************************************************************************************************************************************************
#Region "*************START*************DISPLAY PRODUCT LISTING SAMPLES********************************************************************************************"
    '*****************************************************************************************************************************************************************
#Region "Sample Product Listing Display"

    Public Sub SetSampleProductInfo(title As String, description As String)
        lblWooSampleTitle.Text = title
        txtWooLongDescriptionPreview.Text = description
        wbWooLongDescriptionPreview.DocumentText = description
    End Sub

#End Region


#End Region '---------END---------------BUTTON EVENT HANDLERS---------------------------------------------------------------------------------------------------------
End Class