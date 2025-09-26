Imports System.Data.SqlClient
Imports System.Windows.Forms

Public Class formListings
    Private isLoading As Boolean = False

    Private Sub formListings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        isLoading = True
        Try
            Using conn = DbConnectionManager.CreateOpenConnection()
                Dim cmd As New SqlCommand("SELECT PK_ManufacturerId, ManufacturerName FROM ModelManufacturers ORDER BY ManufacturerName", conn)
                Dim dt As New DataTable()
                dt.Load(cmd.ExecuteReader())
                cmbManufacturerName.DataSource = dt
                cmbManufacturerName.DisplayMember = "ManufacturerName"
                cmbManufacturerName.ValueMember = "PK_ManufacturerId"
                cmbManufacturerName.SelectedIndex = -1
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading manufacturers: " & ex.Message)
        Finally
            isLoading = False
        End Try
    End Sub

    ' Purpose: Loads all series for the selected manufacturer, including FK_equipmentTypeId, and binds to cmbSeriesName.
    ' Dependencies: Imports System.Data.SqlClient, DbConnectionManager
    ' Current date: 2025-09-25
    Private Sub cmbManufacturerName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbManufacturerName.SelectedIndexChanged
        ' >>> changed
        If cmbManufacturerName.SelectedIndex = -1 OrElse cmbManufacturerName.SelectedValue Is Nothing Then
            cmbSeriesName.DataSource = Nothing
            Return
        End If

        Try
            Using conn = DbConnectionManager.CreateOpenConnection()
                ' Now also select FK_equipmentTypeId
                Dim sql As String = "SELECT PK_SeriesId, SeriesName, FK_equipmentTypeId FROM ModelSeries WHERE FK_ManufacturerId = @ManufacturerId ORDER BY SeriesName"
                Using cmd As New SqlCommand(sql, conn)
                    Dim manufacturerId As Object = cmbManufacturerName.SelectedValue
                    If TypeOf manufacturerId Is DataRowView Then
                        manufacturerId = CType(manufacturerId, DataRowView)("PK_ManufacturerId")
                    End If
                    cmd.Parameters.AddWithValue("@ManufacturerId", manufacturerId)
                    Using reader = cmd.ExecuteReader()
                        Dim dt As New DataTable()
                        dt.Load(reader)
                        If Not isLoading AndAlso dt.Rows.Count = 0 Then
                            MessageBox.Show("No series found for this manufacturer.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                        cmbSeriesName.DataSource = dt
                        cmbSeriesName.DisplayMember = "SeriesName"
                        cmbSeriesName.ValueMember = "PK_SeriesId"
                        cmbSeriesName.SelectedIndex = -1
                        ' FK_equipmentTypeId is now available in the DataSource for later use
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading series: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        ' <<< end changed
    End Sub


    ' Purpose: When a series is selected, store the equipmentTypeId for the series and load all models for that series and their latest pricing into dgvListingInformation.
    ' Dependencies: Imports System.Data.SqlClient, DbConnectionManager
    ' Current date: 2025-09-25
    Private currentEquipmentTypeId As Integer = 0 ' >>> changed

    Private Sub cmbSeriesName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSeriesName.SelectedIndexChanged
        ' >>> changed
        ' Only proceed if a series is selected
        If cmbSeriesName.SelectedIndex = -1 OrElse cmbSeriesName.SelectedValue Is Nothing Then
            dgvListingInformation.DataSource = Nothing
            currentEquipmentTypeId = 0 ' Clear equipmentTypeId if nothing selected
            Return
        End If

        ' Get equipmentTypeId from the selected series
        Dim drv As DataRowView = TryCast(cmbSeriesName.SelectedItem, DataRowView)
        If drv IsNot Nothing AndAlso Not IsDBNull(drv("FK_equipmentTypeId")) Then
            currentEquipmentTypeId = Convert.ToInt32(drv("FK_equipmentTypeId")) ' >>> changed
        Else
            currentEquipmentTypeId = 0
        End If

        Try
            Using conn = DbConnectionManager.CreateOpenConnection()
                ' Query: Get all models for the selected series, their equipment type, latest pricing, and parentsku
                Dim sql As String =
"SELECT 
    mfr.ManufacturerName,
    et.equipmentTypeName,
    s.seriesName,
    mo.modelName,
    mo.parentsku, -- Added parentsku column
    pr.RetailPrice_Choice_Amazon,
    pr.RetailPrice_ChoicePadded_Amazon,
    pr.RetailPrice_Leather_Amazon,
    pr.RetailPrice_LeatherPadded_Amazon,
    pr.RetailPrice_Choice_Reverb,
    pr.RetailPrice_ChoicePadded_Reverb,
    pr.RetailPrice_Leather_Reverb,
    pr.RetailPrice_LeatherPadded_Reverb,
    pr.RetailPrice_Choice_eBay,
    pr.RetailPrice_ChoicePadded_eBay,
    pr.RetailPrice_Leather_eBay,
    pr.RetailPrice_LeatherPadded_eBay,
    pr.RetailPrice_Choice_Etsy,
    pr.RetailPrice_ChoicePadded_Etsy,
    pr.RetailPrice_Leather_Etsy,
    pr.RetailPrice_LeatherPadded_Etsy
FROM ModelSeries s
INNER JOIN ModelManufacturers mfr ON s.FK_manufacturerId = mfr.PK_ManufacturerId
INNER JOIN model mo ON mo.FK_seriesId = s.PK_seriesId
INNER JOIN ModelEquipmentTypes et ON s.FK_equipmentTypeId = et.PK_equipmentTypeId
OUTER APPLY (
    SELECT TOP 1 *
    FROM ModelHistoryCostRetailPricing pr
    WHERE pr.FK_ModelId = mo.PK_modelId
    ORDER BY pr.DateCalculated DESC
) pr
WHERE s.PK_seriesId = @SeriesId
ORDER BY mo.modelName"
                Using cmd As New SqlCommand(sql, conn)
                    ' Handle DataRowView issue for SelectedValue
                    Dim seriesId As Object = cmbSeriesName.SelectedValue
                    If TypeOf seriesId Is DataRowView Then
                        seriesId = CType(seriesId, DataRowView)("PK_SeriesId")
                    End If
                    cmd.Parameters.AddWithValue("@SeriesId", seriesId)
                    Using reader = cmd.ExecuteReader()
                        Dim dt As New DataTable()
                        dt.Load(reader)
                        dgvListingInformation.DataSource = dt
                        ' After loading, check for missing parentsku
                        For Each row As DataRow In dt.Rows
                            If row.Table.Columns.Contains("parentsku") Then
                                Dim parentSkuObj = row("parentsku")
                                If parentSkuObj Is DBNull.Value OrElse String.IsNullOrWhiteSpace(parentSkuObj.ToString()) Then
                                    Dim modelName As String = If(row.Table.Columns.Contains("modelName"), row("modelName").ToString(), "(unknown)")
                                    Dim result = MessageBox.Show(
                                        $"Model '{modelName}' does not have a Parent SKU. Would you like to create a Parent SKU now?",
                                        "Missing Parent SKU",
                                        MessageBoxButtons.YesNoCancel,
                                        MessageBoxIcon.Question
                                    )
                                    ' We'll wire up the Yes/No/Cancel logic later as requested
                                    Exit For ' Only prompt once per load
                                End If
                            End If
                        Next
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading model/pricing info: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        ' <<< end changed
    End Sub


    ' Purpose: Handles the click event for btnApiTester, calls Reverb API asynchronously and displays user info.
    ' Dependencies: Imports System.Net.Http, Imports Newtonsoft.Json, ReverbApiClient
    ' Current date: 2025-09-25
    Private Async Sub btnApiTester_Click(sender As Object, e As EventArgs) Handles btnApiTester.Click
        ' >>> changed
        Try
            Dim client As New ReverbApiClient("2c90ace94932a6ca039b0ece7d5b92be5b913ce74b10c9dce8e52fcc345da9e9")
            Dim userInfoJson = Await client.GetUserInfoAsync()
            MessageBox.Show(userInfoJson, "Reverb User Info")
        Catch ex As Exception
            MessageBox.Show("Failed to contact Reverb: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        ' <<< end changed
    End Sub

    ' Purpose: Builds a ReverbListing object using equipmentTypeId, marketplaceId, and model-specific data, generating the title and description by replacing placeholders in the templates.
    ' Dependencies: Imports System.Data.SqlClient, ReverbListing, ReverbPhoto, Newtonsoft.Json
    ' Current date: 2025-09-25
    Private Function BuildReverbListing(equipmentTypeId As Integer, marketplaceId As Integer, modelData As DataRow) As ReverbListing
        ' >>> changed
        Dim allowedFields As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Dim images As New List(Of ReverbPhoto)()
        Dim descriptionTemplate As String = ""
        Dim description As String = ""
        Dim productType As String = ""
        Dim subcategory1 As String = ""
        Dim titleTemplate As String = ""
        Dim title As String = ""

        Try
            Using conn = DbConnectionManager.CreateOpenConnection()
                ' 1. Get allowed values for the key fields (only where FK_equipmentTypeId matches)
                Dim sqlAllowed = "
                SELECT marketplaceReverbNameValuesAllowed, allowedValue
                FROM MpReverbNameValuesAllowed
                WHERE FK_marketplaceNameId = @MarketplaceId AND FK_equipmentTypeId = @EquipmentTypeId"
                Using cmd As New SqlCommand(sqlAllowed, conn)
                    cmd.Parameters.AddWithValue("@MarketplaceId", marketplaceId)
                    cmd.Parameters.AddWithValue("@EquipmentTypeId", equipmentTypeId)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim field = reader("marketplaceReverbNameValuesAllowed").ToString()
                            Dim value = reader("allowedValue").ToString()
                            allowedFields(field) = value
                        End While
                    End Using
                End Using

                ' 2. Assign the key fields from allowedFields
                If allowedFields.ContainsKey("product_type") Then productType = allowedFields("product_type")
                If allowedFields.ContainsKey("subcategory_1") Then subcategory1 = allowedFields("subcategory_1")
                If allowedFields.ContainsKey("title") Then titleTemplate = allowedFields("title")
                If allowedFields.ContainsKey("description") Then descriptionTemplate = allowedFields("description")

                ' 3. Generate the title by replacing placeholders
                If Not String.IsNullOrWhiteSpace(titleTemplate) Then
                    title = titleTemplate _
                        .Replace("[MANUFACTURER_NAME]", modelData("ManufacturerName").ToString()) _
                        .Replace("[SERIES_NAME]", modelData("seriesName").ToString()) _
                        .Replace("[MODEL_NAME]", modelData("modelName").ToString()) _
                        .Replace("[EQUIPMENT_TYPE]", modelData("equipmentTypeName").ToString())
                Else
                    title = modelData("modelName").ToString()
                End If

                ' 4. Generate the description by replacing placeholders
                If Not String.IsNullOrWhiteSpace(descriptionTemplate) Then
                    description = descriptionTemplate _
                        .Replace("[MANUFACTURER_NAME]", modelData("ManufacturerName").ToString()) _
                        .Replace("[SERIES_NAME]", modelData("seriesName").ToString()) _
                        .Replace("[MODEL_NAME]", modelData("modelName").ToString()) _
                        .Replace("[EQUIPMENT_TYPE]", modelData("equipmentTypeName").ToString())
                Else
                    description = "No description available."
                End If

                ' 5. Get images (Photos) for correct marketplace and equipmentTypeId, ordered by position
                Dim sqlImages = "
                    SELECT position, imageUrl
                    FROM ModelEquipmentTypeImage
                    WHERE FK_marketplaceNameId = @MarketplaceId AND FK_equipmentTypeId = @EquipmentTypeId AND isActive = 1
                    ORDER BY position"
                Using cmd As New SqlCommand(sqlImages, conn)
                    cmd.Parameters.AddWithValue("@MarketplaceId", marketplaceId)
                    cmd.Parameters.AddWithValue("@EquipmentTypeId", equipmentTypeId)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim url As String = reader("imageUrl").ToString()
                            If Not String.IsNullOrWhiteSpace(url) Then
                                images.Add(New ReverbPhoto With {.PhotoUrl = url})
                            End If
                        End While
                    End Using
                End Using
            End Using

            ' 6. Build the listing object, using allowedFields for the key fields
            Dim listing As New ReverbListing With {
                .Title = title, ' >>> changed
                .Condition = If(modelData.Table.Columns.Contains("condition") AndAlso Not IsDBNull(modelData("condition")), modelData("condition").ToString(), "excellent"),
                .Inventory = 1,
                .Sku = modelData("parentsku").ToString(),
                .Make = modelData("ManufacturerName").ToString(),
                .Model = modelData("modelName").ToString(),
                .Description = description, ' >>> changed
                .Price = If(modelData.Table.Columns.Contains("RetailPrice_Choice_Reverb") AndAlso Not IsDBNull(modelData("RetailPrice_Choice_Reverb")), Convert.ToDecimal(modelData("RetailPrice_Choice_Reverb")), 0D),
                .ProductType = productType, ' >>> changed
                .Subcategory1 = subcategory1, ' >>> changed
                .OffersEnabled = True,
                .LocalPickup = False,
                .ShippingProfileName = "",
                .UpcDoesNotApply = True,
                .CountryOfOrigin = "",
                .Photos = images
            }
            Return listing
        Catch ex As Exception
            MessageBox.Show("Error building Reverb listing: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
        ' <<< end changed
    End Function

    ' Purpose: Handles the click event for btnReverbListings. Opens the MpListingsPreview form with a preview of the first Reverb listing in dgvListingInformation.
    ' Dependencies: Imports System.Data.SqlClient, ReverbListing, MpListingsPreview, BuildReverbListing, DataGridView, WinForms
    ' Current date: 2025-09-25
    Private Sub btnReverbListings_Click(sender As Object, e As EventArgs) Handles btnReverbListings.Click
        ' >>> changed
        ' Preview the very first model in dgvListingInformation (if any), no selection required.
        If dgvListingInformation.Rows.Count = 0 OrElse dgvListingInformation.DataSource Is Nothing Then
            MessageBox.Show("No models available to preview.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Try
            ' Get the first DataRow from the DataGridView
            Dim firstRow As DataGridViewRow = dgvListingInformation.Rows(0)
            Dim dataRowView As DataRowView = TryCast(firstRow.DataBoundItem, DataRowView)
            If dataRowView Is Nothing Then
                MessageBox.Show("Could not retrieve model data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            Dim modelData As DataRow = dataRowView.Row

            ' Use the stored equipmentTypeId for all Reverb listing queries
            Dim equipmentTypeId As Integer = currentEquipmentTypeId
            Dim marketplaceId As Integer = GetMarketplaceIdByName("Reverb")
            If marketplaceId = 0 Then
                MessageBox.Show("Could not find marketplace ID for 'Reverb'.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If


            ' Build the ReverbListing object
            Dim listing As ReverbListing = BuildReverbListing(equipmentTypeId, marketplaceId, modelData)
            If listing Is Nothing Then Return

            ' Show the preview form
            Dim previewForm As New MpListingsPreview(listing)
            previewForm.ShowDialog(Me)
        Catch ex As Exception
            MessageBox.Show("Failed to open Reverb listing preview: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        ' <<< end changed
    End Sub

    ' Purpose: Debug version of the MpReverbNameValuesDefault query and dictionary assignment.
    ' Dependencies: Imports System.Data.SqlClient, System.Diagnostics, System.Windows.Forms
    ' Current date: 2025-09-25
    Private Sub DebugDefaultFieldsQuery(conn As SqlConnection, marketplaceId As Integer, equipmentTypeId As Integer)
        ' >>> changed
        Dim defaultFields As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Dim debugOutput As New System.Text.StringBuilder()
        debugOutput.AppendLine("DEBUG: Querying MpReverbNameValuesDefault")
        debugOutput.AppendLine($"marketplaceId={marketplaceId}, equipmentTypeId={equipmentTypeId}")

        Dim sqlDefault = "
        SELECT marketplaceReverbFieldName, defaultValue, FK_equipmentTypeId
        FROM MpReverbNameValuesDefault
        WHERE FK_marketplaceNameId = @MarketplaceId
        ORDER BY 
            CASE WHEN FK_equipmentTypeId = @EquipmentTypeId THEN 0 ELSE 1 END"
        Using cmd As New SqlCommand(sqlDefault, conn)
            cmd.Parameters.AddWithValue("@MarketplaceId", marketplaceId)
            cmd.Parameters.AddWithValue("@EquipmentTypeId", equipmentTypeId)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim field = reader("marketplaceReverbFieldName").ToString()
                    Dim defaultValue = reader("defaultValue").ToString()
                    Dim fkEqId As String = If(IsDBNull(reader("FK_equipmentTypeId")), "NULL", reader("FK_equipmentTypeId").ToString())
                    debugOutput.AppendLine($"Row: field={field}, defaultValue={defaultValue}, FK_equipmentTypeId={fkEqId}")
                    If Not defaultFields.ContainsKey(field) Then
                        defaultFields(field) = defaultValue
                        debugOutput.AppendLine($"  -> Added to defaultFields: {field} = {defaultValue}")
                    Else
                        debugOutput.AppendLine($"  -> Skipped (already set): {field}")
                    End If
                End While
            End Using
        End Using

        debugOutput.AppendLine("Final defaultFields dictionary:")
        For Each kvp In defaultFields
            debugOutput.AppendLine($"  {kvp.Key} = {kvp.Value}")
        Next

        Debug.WriteLine(debugOutput.ToString())
        MessageBox.Show(debugOutput.ToString(), "DEBUG: MpReverbNameValuesDefault Results")
        ' <<< end changed
    End Sub

    ' Purpose: Retrieves the PK_marketplaceNameId for a given marketplace name (e.g., "Reverb").
    ' Dependencies: Imports System.Data.SqlClient
    ' Current date: 2025-09-25
    Private Function GetMarketplaceIdByName(marketplaceName As String) As Integer
        ' >>> changed
        Try
            Using conn = DbConnectionManager.CreateOpenConnection()
                Dim sql As String = "SELECT PK_marketplaceNameId FROM MpName WHERE marketplaceName = @marketplaceName"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@marketplaceName", marketplaceName)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        Return Convert.ToInt32(result)
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error retrieving marketplace ID: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return 0
        ' <<< end changed
    End Function
End Class