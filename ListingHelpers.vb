Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports ReverbCode.ReverbListingModel


''' <summary>
''' Provides helper methods for listing-related database operations.
''' </summary>
Public Class ListingHelpers

    ' Purpose: Creates a listing object for the specified marketplace, mapping default fields to the correct model properties. Sets Sku from parentSku and Model from modelName.
    ' Dependencies: Imports ReverbCode.ReverbListingModel
    ' Current date: 2025-09-26

    Public Shared Function CreateMarketplaceListing(
         marketplaceName As String,
         title As String,
         description As String,
         defaultFields As Dictionary(Of String, String),
         parentSku As String,
         modelName As String,
         baseRetailPriceChoice As Decimal ' >>> changed
     ) As ReverbListing
        ' >>> changed
        ' Map fields for Reverb; extend for other marketplaces as needed
        Dim listing As New ReverbListing() With {
            .Title = title,
            .Description = description
        }
        If marketplaceName = "Reverb" Then
            listing.Condition = If(defaultFields.ContainsKey("condition"), defaultFields("condition"), Nothing)
            listing.Inventory = If(defaultFields.ContainsKey("inventory"), Convert.ToInt32(defaultFields("inventory")), 0)
            listing.Make = If(defaultFields.ContainsKey("make"), defaultFields("make"), Nothing)
            listing.Sku = parentSku
            listing.Model = modelName
            listing.Price = baseRetailPriceChoice ' Always use BASERETAILPRICE_CHOICE for Reverb price
            listing.ProductType = If(defaultFields.ContainsKey("product_type"), defaultFields("product_type"), Nothing)
            listing.Subcategory1 = If(defaultFields.ContainsKey("subcategory_1"), defaultFields("subcategory_1"), Nothing)
            listing.OffersEnabled = If(defaultFields.ContainsKey("offers_enabled"), Convert.ToBoolean(defaultFields("offers_enabled")), False)
            listing.LocalPickup = If(defaultFields.ContainsKey("local_pickup"), Convert.ToBoolean(defaultFields("local_pickup")), False)
            listing.ShippingProfileName = If(defaultFields.ContainsKey("shipping_profile_name"), defaultFields("shipping_profile_name"), Nothing)
            listing.UpcDoesNotApply = If(defaultFields.ContainsKey("upc_does_not_apply"), Convert.ToBoolean(defaultFields("upc_does_not_apply")), False)
            listing.CountryOfOrigin = If(defaultFields.ContainsKey("country_of_origin"), defaultFields("country_of_origin"), Nothing)
        End If
        ' <<< end changed
        Return listing
    End Function

    ' Purpose: Processes placeholders in the description template and replaces them with correct pricing and feature values.
    ' Dependencies: Imports System.Data.SqlClient, Imports System.Text.RegularExpressions
    ' Current date: 2025-09-26
    Public Shared Function BuildListingDescription(descriptionTemplate As String, modelId As Integer, marketplaceName As String, equipmentTypeId As Integer) As String
        Dim result As String = descriptionTemplate

        ' Map marketplace names to pricing columns
        Dim priceColumns As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
        {"Reverb", "RetailPrice_Choice_Reverb"},
        {"Amazon", "RetailPrice_Choice_Amazon"},
        {"eBay", "RetailPrice_Choice_eBay"},
        {"Etsy", "RetailPrice_Choice_Etsy"}
    }
        Dim paddedPriceColumns As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
        {"Reverb", "RetailPrice_ChoicePadded_Reverb"},
        {"Amazon", "RetailPrice_ChoicePadded_Amazon"},
        {"eBay", "RetailPrice_ChoicePadded_eBay"},
        {"Etsy", "RetailPrice_ChoicePadded_Etsy"}
    }
        Dim leatherPriceColumns As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
        {"Reverb", "RetailPrice_Leather_Reverb"},
        {"Amazon", "RetailPrice_Leather_Amazon"},
        {"eBay", "RetailPrice_Leather_eBay"},
        {"Etsy", "RetailPrice_Leather_Etsy"}
    }
        Dim leatherPaddedPriceColumns As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
        {"Reverb", "RetailPrice_LeatherPadded_Reverb"},
        {"Amazon", "RetailPrice_LeatherPadded_Amazon"},
        {"eBay", "RetailPrice_LeatherPadded_eBay"},
        {"Etsy", "RetailPrice_LeatherPadded_Etsy"}
    }

        ' Get most recent pricing for this model
        Dim priceValues As New Dictionary(Of String, Decimal)(StringComparer.OrdinalIgnoreCase)
        Try
            Using conn = DbConnectionManager.CreateOpenConnection()
                Dim sql As String =
"SELECT TOP 1
    RetailPrice_Choice_Reverb, RetailPrice_ChoicePadded_Reverb, RetailPrice_Leather_Reverb, RetailPrice_LeatherPadded_Reverb,
    RetailPrice_Choice_Amazon, RetailPrice_ChoicePadded_Amazon, RetailPrice_Leather_Amazon, RetailPrice_LeatherPadded_Amazon,
    RetailPrice_Choice_eBay, RetailPrice_ChoicePadded_eBay, RetailPrice_Leather_eBay, RetailPrice_LeatherPadded_eBay,
    RetailPrice_Choice_Etsy, RetailPrice_ChoicePadded_Etsy, RetailPrice_Leather_Etsy, RetailPrice_LeatherPadded_Etsy
FROM ModelHistoryCostRetailPricing
WHERE FK_ModelId = @ModelId
ORDER BY DateCalculated DESC"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@ModelId", modelId)
                    Using reader = cmd.ExecuteReader()
                        If reader.Read() Then
                            For i As Integer = 0 To reader.FieldCount - 1
                                Dim colName = reader.GetName(i)
                                Dim colValue = If(IsDBNull(reader(i)), "NULL", reader(i).ToString())
                                Console.WriteLine($"DEBUG: Column '{colName}' Value '{colValue}'")
                            Next

                            For Each col In priceColumns.Values.Concat(paddedPriceColumns.Values).Concat(leatherPriceColumns.Values).Concat(leatherPaddedPriceColumns.Values)
                                If Not IsDBNull(reader(col)) Then
                                    priceValues(col) = Convert.ToDecimal(reader(col))
                                    Console.WriteLine($"DEBUG: Loaded {col} = {priceValues(col)}")
                                Else
                                    Console.WriteLine($"DEBUG: {col} is NULL")
                                End If
                            Next
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Console.WriteLine("Error loading pricing: " & ex.Message)
        End Try

        ' Use a composite key: "FEATURE|EQUIPMENTTYPEID"
        Dim featurePrices As New Dictionary(Of String, Decimal)(StringComparer.OrdinalIgnoreCase)
        Try
            Using conn = DbConnectionManager.CreateOpenConnection()
                Dim sql As String = "SELECT DesignFeaturesName, FK_EquipmentTypeId, AddedPrice FROM ModelDesignFeatures WHERE FK_EquipmentTypeId = @EquipmentTypeId"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@EquipmentTypeId", equipmentTypeId)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim name = reader("DesignFeaturesName").ToString().Trim().ToUpperInvariant()
                            Dim eqTypeId = Convert.ToInt32(reader("FK_EquipmentTypeId"))
                            Dim price = If(IsDBNull(reader("AddedPrice")), 0D, Convert.ToDecimal(reader("AddedPrice")))
                            Dim key = $"{name}|{eqTypeId}"
                            featurePrices(key) = price
                            Console.WriteLine($"DEBUG: Loaded feature '{key}' = {price}")
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Console.WriteLine("Error loading feature prices: " & ex.Message)
        End Try

        ' Helper to get price for a placeholder
        Dim functionMap As New Dictionary(Of String, Func(Of String, Decimal)) From {
        {"BASERETAILPRICE_CHOICE", Function(marketplace)
                                       Dim col = If(priceColumns.ContainsKey(marketplace), priceColumns(marketplace), "")
                                       Return If(col <> "" AndAlso priceValues.ContainsKey(col), priceValues(col), 0D)
                                   End Function},
        {"BASERETAILPRICE_CHOICE w/ PADDING", Function(marketplace)
                                                  Dim col = If(paddedPriceColumns.ContainsKey(marketplace), paddedPriceColumns(marketplace), "")
                                                  Return If(col <> "" AndAlso priceValues.ContainsKey(col), priceValues(col), 0D)
                                              End Function},
        {"BASERETAILPRICE_SYNTHETICLEATHER", Function(marketplace)
                                                 Dim col = If(leatherPriceColumns.ContainsKey(marketplace), leatherPriceColumns(marketplace), "")
                                                 Return If(col <> "" AndAlso priceValues.ContainsKey(col), priceValues(col), 0D)
                                             End Function},
        {"BASERETAILPRICE_SYNTHETICLEATHER w/ PADDING", Function(marketplace)
                                                            Dim col = If(leatherPaddedPriceColumns.ContainsKey(marketplace), leatherPaddedPriceColumns(marketplace), "")
                                                            Return If(col <> "" AndAlso priceValues.ContainsKey(col), priceValues(col), 0D)
                                                        End Function}
    }

        Dim pattern As String = "\{\{([A-Z0-9_ ]+(?:\s*\+\s*[A-Z0-9_ ]+)*)\}\}"
        ' >>> changed
        result = Regex.Replace(result, pattern, Function(m)
                                                    Dim parts = m.Groups(1).Value.Split({"+"}, StringSplitOptions.RemoveEmptyEntries)
                                                    Dim total As Decimal = 0D
                                                    Console.WriteLine("DEBUG: Processing placeholder group: " & m.Groups(1).Value)
                                                    For Each part In parts
                                                        Dim key = part.Trim()
                                                        Console.WriteLine($"DEBUG: Placeholder key '{key}'")
                                                        If functionMap.ContainsKey(key) Then
                                                            Dim price = functionMap(key)(marketplaceName)
                                                            Console.WriteLine($"DEBUG: functionMap found '{key}' for marketplace '{marketplaceName}' = {price}")
                                                            total += price
                                                        Else
                                                            Dim compositeKey = $"{key.ToUpperInvariant()}|{equipmentTypeId}"
                                                            If featurePrices.ContainsKey(compositeKey) Then
                                                                Dim featurePrice = featurePrices(compositeKey)
                                                                Console.WriteLine($"DEBUG: featurePrices found '{compositeKey}' = {featurePrice}")
                                                                total += featurePrice
                                                            Else
                                                                Console.WriteLine($"DEBUG: No match for '{compositeKey}' in featurePrices")
                                                            End If
                                                        End If
                                                    Next
                                                    Console.WriteLine($"DEBUG: Total for placeholder '{{{{{m.Groups(1).Value}}}}}' = {total.ToString("C2")}")
                                                    Return total.ToString("C2")
                                                End Function)
        ' <<< end changed

        Return result
    End Function
    ' Purpose: Gets all field names for a marketplace from MpFieldDefinitions.
    ' Dependencies: Imports System.Data.SqlClient
    ' Current date: 2025-09-26
    Public Shared Function GetMarketplaceFieldNames(marketplaceId As Integer) As List(Of String)
        ' >>> changed
        Dim fieldNames As New List(Of String)
        Try
            Using conn = DbConnectionManager.CreateOpenConnection()
                Dim sql As String = "SELECT mpFieldName FROM MpFieldDefinitions WHERE FK_mpNameId = @MarketplaceId ORDER BY mpFieldName"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@MarketplaceId", marketplaceId)
                    Using reader = cmd.ExecuteReader()
                        While reader.Read()
                            fieldNames.Add(reader("mpFieldName").ToString())
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Throw New Exception("Error retrieving field names: " & ex.Message, ex)
        End Try
        Return fieldNames
        ' <<< end changed
    End Function

    ''' <summary>
    ''' Gets default values for all field names for a marketplace and equipment type from MpFieldValues.
    ''' If FK_equipmentTypeId is NULL, use that value. If FK_equipmentTypeId matches, use that value instead.
    ''' </summary>
    ''' <param name="marketplaceId">Marketplace ID</param>
    ''' <param name="equipmentTypeId">Equipment Type ID</param>
    ''' <returns>Dictionary of field name to default value</returns>
    ''' <remarks>
    ''' Dependencies: Imports System.Data.SqlClient
    ''' Current date: 2025-09-26
    ''' </remarks>
    Public Shared Function GetMarketplaceDefaultFieldValues(marketplaceId As Integer, equipmentTypeId As Integer) As Dictionary(Of String, String)
        ' >>> changed
        Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Try
            Using conn = DbConnectionManager.CreateOpenConnection()
                ' Get all field names for this marketplace
                Dim fieldNames = GetMarketplaceFieldNames(marketplaceId)
                If fieldNames.Count = 0 Then Return result

                ' Build a parameterized IN clause for field names
                Dim fieldParams As New List(Of String)
                For i As Integer = 0 To fieldNames.Count - 1
                    fieldParams.Add("@FieldName" & i)
                Next

                ' Corrected SQL: Join MpFieldValues to MpFieldDefinitions to get field names
                Dim sql As String =
"SELECT d.mpFieldName, v.defaultValue, v.FK_equipmentTypeId
 FROM MpFieldValues v
 INNER JOIN MpFieldDefinitions d ON v.FK_mpFieldDefinitionsId = d.PK_mpFieldDefinitionsId
 WHERE v.FK_mpNameId = @MarketplaceId AND d.mpFieldName IN (" & String.Join(",", fieldParams) & ")
 ORDER BY d.mpFieldName, CASE WHEN v.FK_equipmentTypeId = @EquipmentTypeId THEN 0 ELSE 1 END"

                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@MarketplaceId", marketplaceId)
                    cmd.Parameters.AddWithValue("@EquipmentTypeId", equipmentTypeId)
                    For i As Integer = 0 To fieldNames.Count - 1
                        cmd.Parameters.AddWithValue(fieldParams(i), fieldNames(i))
                    Next

                    Using reader = cmd.ExecuteReader()
                        Dim seenFields As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                        While reader.Read()
                            Dim fieldName As String = reader("mpFieldName").ToString()
                            Dim defaultValue As String = reader("defaultValue").ToString()
                            Dim fkEqIdObj = reader("FK_equipmentTypeId")
                            Dim eqTypeMatch As Boolean = Not IsDBNull(fkEqIdObj) AndAlso Convert.ToInt32(fkEqIdObj) = equipmentTypeId
                            Dim eqTypeIsNull As Boolean = IsDBNull(fkEqIdObj)
                            If Not seenFields.Contains(fieldName) Then
                                If eqTypeMatch OrElse eqTypeIsNull Then
                                    result(fieldName) = defaultValue
                                    seenFields.Add(fieldName)
                                End If
                            End If
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Throw New Exception("Error retrieving default field values: " & ex.Message, ex)
        End Try
        Return result
        ' <<< end changed
    End Function

    ' Purpose: Maps image field names (from MpFieldDefinitions) to image URLs (from ModelEquipmentTypeImage) for a marketplace and equipment type.
    ' Dependencies: Imports System.Data.SqlClient
    ' Current date: 2025-09-26
    Public Shared Function GetMarketplaceImageFieldValues(marketplaceId As Integer, equipmentTypeId As Integer) As Dictionary(Of String, String)
        ' >>> changed
        Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Try
            Using conn = DbConnectionManager.CreateOpenConnection()
                ' 1. Get image field names for the marketplace
                Dim fieldNames As New List(Of String)
                Dim sqlFields As String = "SELECT mpFieldName FROM MpFieldDefinitions WHERE FK_mpNameId = @MarketplaceId AND mpFieldName LIKE 'product_image_%' ORDER BY mpFieldName"
                Using cmdFields As New SqlCommand(sqlFields, conn)
                    cmdFields.Parameters.AddWithValue("@MarketplaceId", marketplaceId)
                    Using reader = cmdFields.ExecuteReader()
                        While reader.Read()
                            fieldNames.Add(reader("mpFieldName").ToString())
                        End While
                    End Using
                End Using

                ' 2. Get image URLs for the marketplace and equipment type
                Dim imageUrls As New List(Of String)
                Dim sqlImages As String = "SELECT imageUrl FROM ModelEquipmentTypeImage WHERE FK_mpNameId = @MarketplaceId AND FK_equipmentTypeId = @EquipmentTypeId AND isActive = 1 ORDER BY position"
                Using cmdImages As New SqlCommand(sqlImages, conn)
                    cmdImages.Parameters.AddWithValue("@MarketplaceId", marketplaceId)
                    cmdImages.Parameters.AddWithValue("@EquipmentTypeId", equipmentTypeId)
                    Using reader = cmdImages.ExecuteReader()
                        While reader.Read()
                            imageUrls.Add(reader("imageUrl").ToString())
                        End While
                    End Using
                End Using

                ' 3. Map field names to image URLs
                For i As Integer = 0 To fieldNames.Count - 1
                    If i < imageUrls.Count Then
                        result(fieldNames(i)) = imageUrls(i)
                    Else
                        result(fieldNames(i)) = "" ' No image for this field
                    End If
                Next
            End Using
        Catch ex As Exception
            Throw New Exception("Error mapping image fields to URLs: " & ex.Message, ex)
        End Try
        Return result
        ' <<< end changed
    End Function

    ' Purpose: Gets all default field values for all selected marketplaces and equipment type.
    ' Returns: Dictionary(Of Integer, Dictionary(Of String, String))
    ' Dependencies: Imports System.Data.SqlClient
    ' Current date: 2025-09-26
    Public Shared Function GetAllMarketplacesDefaultFieldValues(marketplaceIds As List(Of Integer), equipmentTypeId As Integer) As Dictionary(Of Integer, Dictionary(Of String, String))
        ' >>> changed
        Dim allResults As New Dictionary(Of Integer, Dictionary(Of String, String))
        For Each marketplaceId In marketplaceIds
            Dim defaults = GetMarketplaceDefaultFieldValues(marketplaceId, equipmentTypeId)
            allResults(marketplaceId) = defaults
        Next
        Return allResults
        ' <<< end changed
    End Function

    ' Purpose: Gets all image field mappings for all selected marketplaces and equipment type.
    ' Returns: Dictionary(Of Integer, Dictionary(Of String, String))
    ' Dependencies: Imports System.Data.SqlClient
    ' Current date: 2025-09-26
    Public Shared Function GetAllMarketplacesImageFieldValues(marketplaceIds As List(Of Integer), equipmentTypeId As Integer) As Dictionary(Of Integer, Dictionary(Of String, String))
        ' >>> changed
        Dim allResults As New Dictionary(Of Integer, Dictionary(Of String, String))
        For Each marketplaceId In marketplaceIds
            Dim imageMap = GetMarketplaceImageFieldValues(marketplaceId, equipmentTypeId)
            allResults(marketplaceId) = imageMap
        Next
        Return allResults
        ' <<< end changed
    End Function
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
    ' <<< end changed
End Class
