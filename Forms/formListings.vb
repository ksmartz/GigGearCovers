Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports ReverbCode.ReverbListingModel

Public Class formListings
    Private isLoading As Boolean = False
    Private checkedListMarketplaces As CheckedListBox
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
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' >>> changed
        ' Create and configure the CheckedListBox
        checkedListMarketplaces = New CheckedListBox() With {
            .Name = "checkedListMarketplaces",
            .Location = New Drawing.Point(20, 20),
            .Size = New Drawing.Size(200, 140),
            .CheckOnClick = True
        }

        ' Add marketplace items
        checkedListMarketplaces.Items.Add("Reverb")
        checkedListMarketplaces.Items.Add("eBay")
        checkedListMarketplaces.Items.Add("Amazon")
        checkedListMarketplaces.Items.Add("Website")
        checkedListMarketplaces.Items.Add("Etsy")
        checkedListMarketplaces.Items.Add("Walmart")
        checkedListMarketplaces.Items.Add("All Marketplaces")

        ' Add the CheckedListBox to the form
        Me.Controls.Add(checkedListMarketplaces)
        ' <<< end changed
    End Sub
    ' >>> changed
    Public Shared Sub UseReverbListingFieldValues(myReverbListing As ReverbListing)
        Try
            ' Get the dictionary of field names and values from the read-only property
            Dim fieldDict As Dictionary(Of String, Object) = myReverbListing.FieldNameValues

            ' Example: Access the "Title" value
            Dim titleValue As String = fieldDict("Title").ToString()

            ' Example: Output all field names and values
            For Each kvp As KeyValuePair(Of String, Object) In fieldDict
                Console.WriteLine($"{kvp.Key}: {kvp.Value}")
            Next
        Catch ex As Exception
            ' Handle any errors gracefully
            Console.WriteLine("Error accessing ReverbListing field values: " & ex.Message)
        End Try
    End Sub


    ''' <summary>
    ''' Builds a listing description by replacing single-brace placeholders and expressions with correct pricing and design feature values.
    ''' Purpose: Handles {BASERETAILPRICE_CHOICE}, {POCKET}, {ZIPPERHANDLE}, and expressions like {BASERETAILPRICE_CHOICE + POCKET + ZIPPERHANDLE}, including "W/ PADDING" variants.
    ''' Dependencies: Imports System.Data.SqlClient, Imports System.Text.RegularExpressions, DbConnectionManager
    ''' Current date: 2025-09-26
    ''' </summary>
    Public Shared Function BuildListingDescription(descriptionTemplate As String, modelId As Integer, marketplaceName As String, equipmentTypeId As Integer, manufacturerName As String, seriesName As String, modelName As String, equipmentTypeName As String) As String
        ' Map marketplaceName to correct column suffix
        Dim priceSuffix As String = ""
        Select Case marketplaceName.ToLowerInvariant()
            Case "reverb"
                priceSuffix = "_Reverb"
            Case "amazon"
                priceSuffix = "_Amazon"
            Case "ebay"
                priceSuffix = "_eBay"
            Case "etsy"
                priceSuffix = "_Etsy"
            Case Else
                priceSuffix = "_Reverb" ' Default to Reverb
        End Select
        MsgBox(descriptionTemplate)
        ' Strongly-typed variables for price values
        Dim baseRetailPriceChoice As Decimal = 0D
        Dim baseRetailPriceChoicePadded As Decimal = 0D
        Dim baseRetailPriceSyntheticLeather As Decimal = 0D
        Dim baseRetailPriceSyntheticLeatherPadded As Decimal = 0D

        ' Arrays for design feature names and prices
        Dim featureNames() As String = {"ZIPPERHANDLE", "POCKET"}
        Dim featurePrices(1) As Decimal

        Try
            Using conn = DbConnectionManager.CreateOpenConnection()
                ' Get latest pricing for the model
                Dim priceSql As String = $"
            SELECT TOP 1
                RetailPrice_Choice{priceSuffix},
                RetailPrice_ChoicePadded{priceSuffix},
                RetailPrice_Leather{priceSuffix},
                RetailPrice_LeatherPadded{priceSuffix}
            FROM ModelHistoryCostRetailPricing
            WHERE FK_ModelId = @ModelId
            ORDER BY DateCalculated DESC"
                Using cmd As New SqlCommand(priceSql, conn)
                    cmd.Parameters.AddWithValue("@ModelId", modelId)
                    Using reader = cmd.ExecuteReader()
                        If reader.Read() Then
                            baseRetailPriceChoice = If(IsDBNull(reader(0)), 0D, Convert.ToDecimal(reader(0)))
                            baseRetailPriceChoicePadded = If(IsDBNull(reader(1)), 0D, Convert.ToDecimal(reader(1)))
                            baseRetailPriceSyntheticLeather = If(IsDBNull(reader(2)), 0D, Convert.ToDecimal(reader(2)))
                            baseRetailPriceSyntheticLeatherPadded = If(IsDBNull(reader(3)), 0D, Convert.ToDecimal(reader(3)))
                        End If
                    End Using
                End Using

                ' Get design feature prices for current equipment type
                Dim featureSql As String = "
            SELECT DesignFeaturesName, AddedPrice
            FROM ModelDesignFeatures
            WHERE FK_EquipmentTypeId = @EquipmentTypeId"
                Using cmd As New SqlCommand(featureSql, conn)
                    cmd.Parameters.AddWithValue("@EquipmentTypeId", equipmentTypeId)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim name As String = reader("DesignFeaturesName").ToString().Trim().ToUpperInvariant()
                            Dim price As Decimal = If(IsDBNull(reader("AddedPrice")), 0D, Convert.ToDecimal(reader("AddedPrice")))
                            For i As Integer = 0 To featureNames.Length - 1
                                If name = featureNames(i) Then
                                    featurePrices(i) = price
                                End If
                            Next
                        End While
                    End Using
                End Using
            End Using

            ' >>> changed
            ' Normalize double braces to single braces for all templates
            descriptionTemplate = Regex.Replace(descriptionTemplate, "\{\{([^\}]+)\}\}", "{$1}")

            Dim result As String = descriptionTemplate
            Dim exprPattern As String = "\{([^\}]+)\}"
            result = Regex.Replace(result, exprPattern, Function(m)
                                                            Dim expr As String = m.Groups(1).Value
                                                            Dim sum As Decimal = 0D
                                                            Dim isSum As Boolean = False
                                                            Dim parts = expr.Split({"+"}, StringSplitOptions.RemoveEmptyEntries)
                                                            For Each part In parts
                                                                Dim trimmed = part.Trim()
                                                                If trimmed.ToUpperInvariant().Contains("W/ PADDING") Then
                                                                    trimmed = Regex.Replace(trimmed, "w/ padding", "W/ PADDING", RegexOptions.IgnoreCase)
                                                                End If
                                                                Select Case True
                                                                    Case String.Equals(trimmed, "MANUFACTURER_NAME", StringComparison.OrdinalIgnoreCase)
                                                                        If parts.Length = 1 Then Return manufacturerName
                                                                    Case String.Equals(trimmed, "SERIES_NAME", StringComparison.OrdinalIgnoreCase)
                                                                        If parts.Length = 1 Then Return seriesName
                                                                    Case String.Equals(trimmed, "MODEL_NAME", StringComparison.OrdinalIgnoreCase)
                                                                        If parts.Length = 1 Then Return modelName
                                                                    Case String.Equals(trimmed, "EQUIPMENT_TYPE", StringComparison.OrdinalIgnoreCase)
                                                                        If parts.Length = 1 Then Return equipmentTypeName
                                                                    Case String.Equals(trimmed, "BASERETAILPRICE_CHOICE", StringComparison.OrdinalIgnoreCase)
                                                                        sum += baseRetailPriceChoice
                                                                        isSum = True
                                                                    Case String.Equals(trimmed, "BASERETAILPRICE_CHOICE W/ PADDING", StringComparison.OrdinalIgnoreCase)
                                                                        sum += baseRetailPriceChoicePadded
                                                                        isSum = True
                                                                    Case String.Equals(trimmed, "BASERETAILPRICE_SYNTHETICLEATHER", StringComparison.OrdinalIgnoreCase)
                                                                        sum += baseRetailPriceSyntheticLeather
                                                                        isSum = True
                                                                    Case String.Equals(trimmed, "BASERETAILPRICE_SYNTHETICLEATHER W/ PADDING", StringComparison.OrdinalIgnoreCase)
                                                                        sum += baseRetailPriceSyntheticLeatherPadded
                                                                        isSum = True
                                                                    Case String.Equals(trimmed, "ZIPPERHANDLE", StringComparison.OrdinalIgnoreCase)
                                                                        sum += featurePrices(0)
                                                                        isSum = True
                                                                    Case String.Equals(trimmed, "POCKET", StringComparison.OrdinalIgnoreCase)
                                                                        sum += featurePrices(1)
                                                                        isSum = True
                                                                End Select
                                                            Next
                                                            If isSum Then
                                                                Return sum.ToString("C2")
                                                            Else
                                                                Return m.Value
                                                            End If
                                                        End Function)
            ' <<< end changed

            Return result
        Catch ex As Exception
            Return $"Error building description: {ex.Message}"
        End Try
    End Function
    ' Purpose: Retrieves the PK_marketplaceNameId for each selected marketplace in checkedListMarketplaces.
    '          If "All Marketplaces" is selected, returns all marketplace IDs from the database.
    ' Dependencies: Imports System.Data.SqlClient, System.Collections.Generic
    ' Current date: 2025-09-26

    Private Function GetSelectedMarketplaceIds() As List(Of Integer)
        ' >>> changed
        Dim selectedIds As New List(Of Integer)()
        Try
            ' Gather selected marketplace names
            Dim selectedNames As New List(Of String)()
            For i As Integer = 0 To checkedListMarketplaces.CheckedItems.Count - 1
                selectedNames.Add(checkedListMarketplaces.CheckedItems(i).ToString())
            Next

            ' If "All Marketplaces" is selected, fetch all IDs from the database
            If selectedNames.Contains("All Marketplaces") Then
                Using conn = DbConnectionManager.CreateOpenConnection()
                    Dim sql As String = "SELECT PK_marketplaceNameId FROM MpName"
                    Using cmd As New SqlCommand(sql, conn)
                        Using reader = cmd.ExecuteReader()
                            While reader.Read()
                                selectedIds.Add(Convert.ToInt32(reader("PK_marketplaceNameId")))
                            End While
                        End Using
                    End Using
                End Using
            Else
                ' Otherwise, fetch IDs for each selected marketplace name
                Using conn = DbConnectionManager.CreateOpenConnection()
                    For Each marketplaceName In selectedNames
                        If marketplaceName <> "All Marketplaces" Then
                            Dim sql As String = "SELECT PK_marketplaceNameId FROM MpName WHERE marketplaceName = @marketplaceName"
                            Using cmd As New SqlCommand(sql, conn)
                                cmd.Parameters.AddWithValue("@marketplaceName", marketplaceName)
                                Dim result = cmd.ExecuteScalar()
                                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                                    selectedIds.Add(Convert.ToInt32(result))
                                End If
                            End Using
                        End If
                    Next
                End Using
            End If
        Catch ex As Exception
            MessageBox.Show("Error retrieving selected marketplace IDs: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return selectedIds
        ' <<< end changed
    End Function
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

    ''' <summary>
    ''' Handles selection change for series combo box. Loads all models for the selected series and their latest pricing info into dgvListingInformation.
    ''' Purpose: Loads manufacturer, equipment type, series, and latest pricing for all models in the selected series.
    ''' Dependencies: Imports System.Data.SqlClient, DbConnectionManager, System.Windows.Forms
    ''' Current date: 2025-09-26
    ''' </summary>
    Private Sub cmbSeriesName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSeriesName.SelectedIndexChanged
        ' Only proceed if a series is selected
        If cmbSeriesName.SelectedIndex = -1 OrElse cmbSeriesName.SelectedValue Is Nothing Then
            dgvListingInformation.DataSource = Nothing
            currentEquipmentTypeId = 0 ' Clear equipmentTypeId if nothing selected
            lblEquipmentType.Text = "" ' Clear label if nothing selected
            Return
        End If

        ' Get equipmentTypeId from the selected series
        Dim drv As DataRowView = TryCast(cmbSeriesName.SelectedItem, DataRowView)
        If drv IsNot Nothing AndAlso Not IsDBNull(drv("FK_equipmentTypeId")) Then
            currentEquipmentTypeId = Convert.ToInt32(drv("FK_equipmentTypeId"))
        Else
            currentEquipmentTypeId = 0
        End If

        ' Update lblEquipmentType with the equipment type name using DbConnectionManager
        Try
            If currentEquipmentTypeId <> 0 Then
                Using conn = DbConnectionManager.CreateOpenConnection()
                    Dim sqlType As String = "SELECT equipmentTypeName FROM ModelEquipmentTypes WHERE PK_equipmentTypeId = @EquipmentTypeId"
                    Using cmdType As New SqlCommand(sqlType, conn)
                        cmdType.Parameters.AddWithValue("@EquipmentTypeId", currentEquipmentTypeId)
                        Dim typeResult = cmdType.ExecuteScalar()
                        If typeResult IsNot Nothing AndAlso Not IsDBNull(typeResult) Then
                            lblEquipmentType.Text = typeResult.ToString()
                        Else
                            lblEquipmentType.Text = "(Unknown Type)"
                        End If
                    End Using
                End Using
            Else
                lblEquipmentType.Text = "(Unknown Type)"
            End If
        Catch ex As Exception
            lblEquipmentType.Text = "(Error)"
            MessageBox.Show("Error retrieving equipment type name: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        ' Load all models for the selected series using DbConnectionManager
        Try
            Using conn = DbConnectionManager.CreateOpenConnection()
                ' >>> changed
                ' Only include the columns you requested in the SELECT statement.
                Dim sql As String =
"SELECT 
    mfr.ManufacturerName,
    et.equipmentTypeName,
    s.SeriesName,
    mo.ModelName,
    mo.parentSku, -- <<< added
    mo.PK_ModelId AS ModelId, -- >>> added, ensures modelId is available
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
INNER JOIN ModelManufacturers mfr ON s.FK_ManufacturerId = mfr.PK_ManufacturerId
INNER JOIN Model mo ON mo.FK_SeriesId = s.PK_SeriesId
INNER JOIN ModelEquipmentTypes et ON s.FK_equipmentTypeId = et.PK_equipmentTypeId
OUTER APPLY (
    SELECT TOP 1
        RetailPrice_Choice_Amazon,
        RetailPrice_ChoicePadded_Amazon,
        RetailPrice_Leather_Amazon,
        RetailPrice_LeatherPadded_Amazon,
        RetailPrice_Choice_Reverb,
        RetailPrice_ChoicePadded_Reverb,
        RetailPrice_Leather_Reverb,
        RetailPrice_LeatherPadded_Reverb,
        RetailPrice_Choice_eBay,
        RetailPrice_ChoicePadded_eBay,
        RetailPrice_Leather_eBay,
        RetailPrice_LeatherPadded_eBay,
        RetailPrice_Choice_Etsy,
        RetailPrice_ChoicePadded_Etsy,
        RetailPrice_Leather_Etsy,
        RetailPrice_LeatherPadded_Etsy
    FROM ModelHistoryCostRetailPricing pr
    WHERE pr.FK_ModelId = mo.PK_ModelId
    ORDER BY pr.DateCalculated DESC
) pr
WHERE s.PK_SeriesId = @SeriesId
ORDER BY mo.ModelName"
                ' <<< end changed
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
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading model/pricing info: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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

    ' Purpose: Clears all checkboxes in checkedListMarketplaces when btnClearMarketPlaces is clicked.
    ' Dependencies: Imports System.Windows.Forms
    ' Current date: 2025-09-26

    Private Sub btnClearMarketPlaces_Click(sender As Object, e As EventArgs) Handles btnClearMarketPlaces.Click
        ' >>> changed
        ' Loop through all items and uncheck them
        For i As Integer = 0 To checkedListMarketplaces.Items.Count - 1
            checkedListMarketplaces.SetItemChecked(i, False)
        Next
        ' <<< end changed
    End Sub

    Private Sub btnReverbListings_Click(sender As Object, e As EventArgs) Handles btnReverbListings.Click

    End Sub
    ''' <summary>
    ''' Handles Create Listings button click. Opens MpListingsPreview for the first model in the grid and populates description, title, and field values using marketplace-agnostic helpers.
    ''' Purpose: Loads default field values, image field values, and description template from mpFieldValues for the selected marketplace and equipment type.
    ''' Dependencies: Imports System.Data.SqlClient, ListingHelpers, Forms\MpListingsPreview, ReverbListing, Imports System.Windows.Forms
    ''' Current date: 2025-09-26
    ''' </summary>
    Private Sub btnCreateListings_Click(sender As Object, e As EventArgs) Handles btnCreateListings.Click
        ' Validate manufacturer and series selection
        If cmbManufacturerName.SelectedIndex = -1 OrElse cmbManufacturerName.SelectedValue Is Nothing Then
            MessageBox.Show("Please select a manufacturer before creating listings.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        If cmbSeriesName.SelectedIndex = -1 OrElse cmbSeriesName.SelectedValue Is Nothing Then
            MessageBox.Show("Please select a series before creating listings.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        If currentEquipmentTypeId = 0 Then
            MessageBox.Show("Equipment type is not set. Please select a valid series.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Get the first checked marketplace (for testing)
        Dim selectedMarketplaceId As Integer = 0
        If checkedListMarketplaces.CheckedItems.Count = 0 Then
            MessageBox.Show("Please select at least one marketplace.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Dim selectedMarketplaceName As String = checkedListMarketplaces.CheckedItems(0).ToString()
        selectedMarketplaceId = GetMarketplaceIdByName(selectedMarketplaceName)
        If selectedMarketplaceId = 0 Then
            MessageBox.Show("Could not find marketplace ID for '" & selectedMarketplaceName & "'.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Try
            ' Gather default field values
            Dim defaultFields = ListingHelpers.GetMarketplaceDefaultFieldValues(selectedMarketplaceId, currentEquipmentTypeId)
            ' Gather image field values
            Dim imageFields = ListingHelpers.GetMarketplaceImageFieldValues(selectedMarketplaceId, currentEquipmentTypeId)

            ' Get the first model in the grid
            If dgvListingInformation.DataSource IsNot Nothing Then
                Dim dt As DataTable = TryCast(dgvListingInformation.DataSource, DataTable)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Dim row As DataRow = dt.Rows(0)
                    ' Dynamically load the description template from mpFieldValues for the selected marketplace and equipment type
                    Dim descriptionTemplate As String = ""
                    Using conn As SqlConnection = DbConnectionManager.CreateOpenConnection()
                        ' Get the PK_mpFieldDefinitionsId for "description"
                        Dim getFieldDefSql As String = "SELECT PK_mpFieldDefinitionsId FROM mpFieldDefinitions WHERE mpFieldName = @FieldName"
                        Dim fieldDefId As Integer = 0
                        Using cmdFieldDef As New SqlCommand(getFieldDefSql, conn)
                            cmdFieldDef.Parameters.AddWithValue("@FieldName", "description")
                            Dim result = cmdFieldDef.ExecuteScalar()
                            If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                                fieldDefId = Convert.ToInt32(result)
                            End If
                        End Using

                        If fieldDefId <> 0 Then
                            Dim sql As String = "SELECT TOP 1 defaultValue FROM mpFieldValues WHERE FK_mpNameId = @MarketplaceId AND FK_equipmentTypeId = @EquipmentTypeId AND FK_mpFieldDefinitionsId = @FieldDefId"
                            Using cmd As New SqlCommand(sql, conn)
                                cmd.Parameters.AddWithValue("@MarketplaceId", selectedMarketplaceId)
                                cmd.Parameters.AddWithValue("@EquipmentTypeId", currentEquipmentTypeId)
                                cmd.Parameters.AddWithValue("@FieldDefId", fieldDefId)
                                Dim result = cmd.ExecuteScalar()
                                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                                    descriptionTemplate = result.ToString()
                                End If
                            End Using
                        End If
                    End Using
                    If String.IsNullOrWhiteSpace(descriptionTemplate) Then
                        descriptionTemplate = "<li>Base Price: {{BASERETAILPRICE_CHOICE}}</li><li>With Pocket: {{BASERETAILPRICE_CHOICE + POCKET}}</li><li>With Zipper: {{BASERETAILPRICE_CHOICE + ZIPPERHANDLE}}</li><li>All: {{BASERETAILPRICE_CHOICE + POCKET + ZIPPERHANDLE}}</li>"
                    End If

                    Dim modelId As Integer = 0
                    If row.Table.Columns.Contains("ModelId") Then
                        modelId = Convert.ToInt32(row("ModelId"))
                    ElseIf row.Table.Columns.Contains("PK_ModelId") Then
                        modelId = Convert.ToInt32(row("PK_ModelId"))
                    End If

                    ' >>> changed
                    ' Build the description using the local function, not ListingHelpers
                    Dim description As String = BuildListingDescription(descriptionTemplate, modelId, selectedMarketplaceName, currentEquipmentTypeId, row("ManufacturerName").ToString(), row("SeriesName").ToString(), row("ModelName").ToString(), lblEquipmentType.Text)
                    ' <<< end changed

                    ' Show a pop-up with the IDs used for pricing
                    MessageBox.Show(
                $"MarketplaceID: {selectedMarketplaceId}{Environment.NewLine}EquipmentTypeID: {currentEquipmentTypeId}{Environment.NewLine}ModelID: {modelId}",
                "Pricing Source Info",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )

                    ' Build title (example: ManufacturerName SeriesName ModelName EquipmentTypeName)
                    Dim titleParts As New List(Of String)
                    If row.Table.Columns.Contains("ManufacturerName") Then titleParts.Add(row("ManufacturerName").ToString())
                    If row.Table.Columns.Contains("SeriesName") Then titleParts.Add(row("SeriesName").ToString())
                    If row.Table.Columns.Contains("ModelName") Then titleParts.Add(row("ModelName").ToString())
                    If row.Table.Columns.Contains("EquipmentTypeName") Then titleParts.Add(row("EquipmentTypeName").ToString())
                    Dim title As String = String.Join(" ", titleParts)

                    Dim parentSku As String = ""
                    If row.Table.Columns.Contains("parentSku") AndAlso Not IsDBNull(row("parentSku")) Then
                        parentSku = row("parentSku").ToString().Trim()
                    End If

                    Dim modelName As String = ""
                    If row.Table.Columns.Contains("ModelName") AndAlso Not IsDBNull(row("ModelName")) Then
                        modelName = row("ModelName").ToString().Trim()
                    End If

                    Dim baseRetailPriceChoice As Decimal = 0D
                    If selectedMarketplaceName = "Reverb" Then
                        If row.Table.Columns.Contains("RetailPrice_Choice_Reverb") AndAlso Not IsDBNull(row("RetailPrice_Choice_Reverb")) Then
                            baseRetailPriceChoice = Convert.ToDecimal(row("RetailPrice_Choice_Reverb"))
                        End If
                    End If

                    Dim listing As ReverbListing = ListingHelpers.CreateMarketplaceListing(
                selectedMarketplaceName,
                title,
                description,
                defaultFields,
                parentSku,
                modelName,
                baseRetailPriceChoice
            )

                    Dim previewForm As New MpListingsPreview(listing)
                    previewForm.ShowDialog()
                Else
                    MessageBox.Show("No models found for the selected series.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Else
                MessageBox.Show("No model data loaded. Please select a series.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("Error creating listings: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class