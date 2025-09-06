Module HelperFunctions

    Public Function CalculateRetailPrice(grandTotal As Decimal, profit As Decimal) As Decimal
        Dim price As Decimal = grandTotal + profit
        ' Round up to the next .95
        Dim rounded As Decimal = Math.Ceiling(price - 0.95D) + 0.95D
        Return rounded
    End Function

    Public Function CalculateFabricWeightInOunces(weightPerLinearYard As Decimal, fabricRollWidth As Decimal, totalFabricSquareInches As Decimal) As Decimal
        If fabricRollWidth <= 0 Then Return 0D
        Dim weightPerSqInch As Decimal = weightPerLinearYard / (36D * fabricRollWidth)
        Return Math.Round(weightPerSqInch * totalFabricSquareInches, 2)
    End Function
    ' In HelperFunctions.vb or frmAddModelInformation.vb

    Public Function CalculateModelMaterialWeights(totalFabricSquareInches As Decimal) As (
    weight_PaddingOnly As Decimal?,
    weight_ChoiceWaterproof As Decimal?,
    weight_ChoiceWaterproof_Padded As Decimal?,
    weight_PremiumSyntheticLeather As Decimal?,
    weight_PremiumSyntheticLeather_Padded As Decimal?
)
        Dim db As New DbConnectionManager()

        Dim rowCW = db.GetActiveFabricBrandProductName("Choice Waterproof")
        Dim rowSL = db.GetActiveFabricBrandProductName("Premium Synthetic Leather")
        Dim rowPad = db.GetActiveFabricBrandProductName("Padding")

        Dim weightCW As Decimal? = Nothing
        Dim weightSL As Decimal? = Nothing
        Dim weightPad As Decimal? = Nothing
        Dim weightCWPad As Decimal? = Nothing
        Dim weightSLPad As Decimal? = Nothing

        If rowCW IsNot Nothing Then
            weightCW = CalculateFabricWeightInOunces(
            Convert.ToDecimal(rowCW("WeightPerLinearYard")),
            Convert.ToDecimal(rowCW("FabricRollWidth")),
            totalFabricSquareInches)
        End If
        If rowSL IsNot Nothing Then
            weightSL = CalculateFabricWeightInOunces(
            Convert.ToDecimal(rowSL("WeightPerLinearYard")),
            Convert.ToDecimal(rowSL("FabricRollWidth")),
            totalFabricSquareInches)
        End If
        If rowPad IsNot Nothing Then
            weightPad = CalculateFabricWeightInOunces(
            Convert.ToDecimal(rowPad("WeightPerLinearYard")),
            Convert.ToDecimal(rowPad("FabricRollWidth")),
            totalFabricSquareInches)
        End If

        If weightCW.HasValue AndAlso weightPad.HasValue Then
            weightCWPad = Math.Round(weightCW.Value + weightPad.Value, 2)
        End If
        If weightSL.HasValue AndAlso weightPad.HasValue Then
            weightSLPad = Math.Round(weightSL.Value + weightPad.Value, 2)
        End If

        Return (weightPad, weightCW, weightCWPad, weightSL, weightSLPad)
    End Function
    ' Place this in your form or a shared module as needed
    Public Sub CalculateAndInsertModelLaborCosts(
    modelId As Integer,
    totalFabricSquareInches As Decimal,
    costs As (costPerSqInch_ChoiceWaterproof As Decimal?, costPerSqInch_PremiumSyntheticLeather As Decimal?, costPerSqInch_Padding As Decimal?, baseCost_ChoiceWaterproof As Decimal?, baseCost_PremiumSyntheticLeather As Decimal?, baseCost_ChoiceWaterproof_Padded As Decimal?, baseCost_PremiumSyntheticLeather_Padded As Decimal?, baseCost_PaddingOnly As Decimal?),
    weights As (weight_PaddingOnly As Decimal?, weight_ChoiceWaterproof As Decimal?, weight_ChoiceWaterproof_Padded As Decimal?, weight_PremiumSyntheticLeather As Decimal?, weight_PremiumSyntheticLeather_Padded As Decimal?),
    notes As String,
    profit_Choice As Decimal?,
    profit_ChoicePadded As Decimal?,
    profit_Leather As Decimal?,
    profit_LeatherPadded As Decimal?,
    AmazonFee_Choice As Decimal?,
    AmazonFee_ChoicePadded As Decimal?,
    AmazonFee_Leather As Decimal?,
    AmazonFee_LeatherPadded As Decimal?,
    ReverbFee_Choice As Decimal?,
    ReverbFee_ChoicePadded As Decimal?,
    ReverbFee_Leather As Decimal?,
    ReverbFee_LeatherPadded As Decimal?,
    BaseCost_GrandTotal_Choice_Amazon As Decimal?,
    BaseCost_GrandTotal_ChoicePadded_Amazon As Decimal?,
    BaseCost_GrandTotal_Leather_Amazon As Decimal?,
    BaseCost_GrandTotal_LeatherPadded_Amazon As Decimal?,
    BaseCost_GrandTotal_Choice_Reverb As Decimal?,
    BaseCost_GrandTotal_ChoicePadded_Reverb As Decimal?,
    BaseCost_GrandTotal_Leather_Reverb As Decimal?,
    BaseCost_GrandTotal_LeatherPadded_Reverb As Decimal?,
    BaseCost_GrandTotal_Choice_eBay As Decimal?,
    BaseCost_GrandTotal_ChoicePadded_eBay As Decimal?,
    BaseCost_GrandTotal_Leather_eBay As Decimal?,
    BaseCost_GrandTotal_LeatherPadded_eBay As Decimal?,
    BaseCost_GrandTotal_Choice_Etsy As Decimal?,
    BaseCost_GrandTotal_ChoicePadded_Etsy As Decimal?,
    BaseCost_GrandTotal_Leather_Etsy As Decimal?,
    BaseCost_GrandTotal_LeatherPadded_Etsy As Decimal?
)
        Dim db As New DbConnectionManager()
        Dim wastePercent As Decimal = 5D

        ' Calculate shipping cost for each fabric type/combination (shipping only)
        Dim shipping_Choice As Decimal = db.GetShippingCostByWeight(If(weights.weight_ChoiceWaterproof, 0D))
        Dim shipping_ChoicePadded As Decimal = db.GetShippingCostByWeight(If(weights.weight_ChoiceWaterproof_Padded, 0D))
        Dim shipping_Leather As Decimal = db.GetShippingCostByWeight(If(weights.weight_PremiumSyntheticLeather, 0D))
        Dim shipping_LeatherPadded As Decimal = db.GetShippingCostByWeight(If(weights.weight_PremiumSyntheticLeather_Padded, 0D))

        ' Calculate base fabric cost + shipping for each type
        Dim baseFabricCost_Choice_Weight As Decimal? = If(costs.baseCost_ChoiceWaterproof, 0D) + shipping_Choice
        Dim baseFabricCost_ChoicePadding_Weight As Decimal? = If(costs.baseCost_ChoiceWaterproof_Padded, 0D) + shipping_ChoicePadded
        Dim baseFabricCost_Leather_Weight As Decimal? = If(costs.baseCost_PremiumSyntheticLeather, 0D) + shipping_Leather
        Dim baseFabricCost_LeatherPadding_Weight As Decimal? = If(costs.baseCost_PremiumSyntheticLeather_Padded, 0D) + shipping_LeatherPadded

        ' --- Labor cost calculations ---
        Dim hourlyRate As Decimal = 17D
        Dim CostLaborNoPadding As Decimal = 0.5D * hourlyRate
        Dim CostLaborWithPadding As Decimal = 1D * hourlyRate

        Dim BaseCost_Choice_Labor As Decimal? = If(baseFabricCost_Choice_Weight, 0D) + CostLaborNoPadding
        Dim BaseCost_ChoicePadding_Labor As Decimal? = If(baseFabricCost_ChoicePadding_Weight, 0D) + CostLaborWithPadding
        Dim BaseCost_Leather_Labor As Decimal? = If(baseFabricCost_Leather_Weight, 0D) + CostLaborNoPadding
        Dim BaseCost_LeatherPadding_Labor As Decimal? = If(baseFabricCost_LeatherPadding_Weight, 0D) + CostLaborWithPadding

        ' --- Marketplace Fee Calculations for eBay and Etsy ---
        Dim eBayFeePercent As Decimal = db.GetMarketplaceFeePercentage("eBay") / 100D
        Dim etsyFeePercent As Decimal = db.GetMarketplaceFeePercentage("Etsy") / 100D

        Dim eBayFee_Choice As Decimal = If(BaseCost_GrandTotal_Choice_eBay, 0D) * eBayFeePercent
        Dim eBayFee_ChoicePadded As Decimal = If(BaseCost_GrandTotal_ChoicePadded_eBay, 0D) * eBayFeePercent
        Dim eBayFee_Leather As Decimal = If(BaseCost_GrandTotal_Leather_eBay, 0D) * eBayFeePercent
        Dim eBayFee_LeatherPadded As Decimal = If(BaseCost_GrandTotal_LeatherPadded_eBay, 0D) * eBayFeePercent

        Dim EtsyFee_Choice As Decimal = If(BaseCost_GrandTotal_Choice_Etsy, 0D) * etsyFeePercent
        Dim EtsyFee_ChoicePadded As Decimal = If(BaseCost_GrandTotal_ChoicePadded_Etsy, 0D) * etsyFeePercent
        Dim EtsyFee_Leather As Decimal = If(BaseCost_GrandTotal_Leather_Etsy, 0D) * etsyFeePercent
        Dim EtsyFee_LeatherPadded As Decimal = If(BaseCost_GrandTotal_LeatherPadded_Etsy, 0D) * etsyFeePercent

        ' --- Retail Price Placeholders (replace with your actual calculations) ---
        Dim RetailPrice_Choice_Amazon As Decimal = 0D
        Dim RetailPrice_ChoicePadded_Amazon As Decimal = 0D
        Dim RetailPrice_Leather_Amazon As Decimal = 0D
        Dim RetailPrice_LeatherPadded_Amazon As Decimal = 0D
        Dim RetailPrice_Choice_Reverb As Decimal = 0D
        Dim RetailPrice_ChoicePadded_Reverb As Decimal = 0D
        Dim RetailPrice_Leather_Reverb As Decimal = 0D
        Dim RetailPrice_LeatherPadded_Reverb As Decimal = 0D
        Dim RetailPrice_Choice_eBay As Decimal = 0D
        Dim RetailPrice_ChoicePadded_eBay As Decimal = 0D
        Dim RetailPrice_Leather_eBay As Decimal = 0D
        Dim RetailPrice_LeatherPadded_eBay As Decimal = 0D
        Dim RetailPrice_Choice_Etsy As Decimal = 0D
        Dim RetailPrice_ChoicePadded_Etsy As Decimal = 0D
        Dim RetailPrice_Leather_Etsy As Decimal = 0D
        Dim RetailPrice_LeatherPadded_Etsy As Decimal = 0D

        ' Insert into history table (all arguments, in order)
        db.InsertModelHistoryCostRetailPricing(
        modelId,
        costs.costPerSqInch_ChoiceWaterproof,
        costs.costPerSqInch_PremiumSyntheticLeather,
        costs.costPerSqInch_Padding,
        totalFabricSquareInches,
        wastePercent,
        costs.baseCost_ChoiceWaterproof,
        costs.baseCost_PremiumSyntheticLeather,
        costs.baseCost_ChoiceWaterproof_Padded,
        costs.baseCost_PremiumSyntheticLeather_Padded,
        costs.baseCost_PaddingOnly,
        weights.weight_PaddingOnly,
        weights.weight_ChoiceWaterproof,
        weights.weight_ChoiceWaterproof_Padded,
        weights.weight_PremiumSyntheticLeather,
        weights.weight_PremiumSyntheticLeather_Padded,
        shipping_Choice,
        shipping_ChoicePadded,
        shipping_Leather,
        shipping_LeatherPadded,
        baseFabricCost_Choice_Weight,
        baseFabricCost_ChoicePadding_Weight,
        baseFabricCost_Leather_Weight,
        baseFabricCost_LeatherPadding_Weight,
        BaseCost_Choice_Labor,
        BaseCost_ChoicePadding_Labor,
        BaseCost_Leather_Labor,
        BaseCost_LeatherPadding_Labor,
        If(profit_Choice, 0D),
        If(profit_ChoicePadded, 0D),
        If(profit_Leather, 0D),
        If(profit_LeatherPadded, 0D),
        If(AmazonFee_Choice, 0D),
        If(AmazonFee_ChoicePadded, 0D),
        If(AmazonFee_Leather, 0D),
        If(AmazonFee_LeatherPadded, 0D),
        If(ReverbFee_Choice, 0D),
        If(ReverbFee_ChoicePadded, 0D),
        If(ReverbFee_Leather, 0D),
        If(ReverbFee_LeatherPadded, 0D),
        eBayFee_Choice,
        eBayFee_ChoicePadded,
        eBayFee_Leather,
        eBayFee_LeatherPadded,
        EtsyFee_Choice,
        EtsyFee_ChoicePadded,
        EtsyFee_Leather,
        EtsyFee_LeatherPadded,
        If(BaseCost_GrandTotal_Choice_Amazon, 0D),
        If(BaseCost_GrandTotal_ChoicePadded_Amazon, 0D),
        If(BaseCost_GrandTotal_Leather_Amazon, 0D),
        If(BaseCost_GrandTotal_LeatherPadded_Amazon, 0D),
        If(BaseCost_GrandTotal_Choice_Reverb, 0D),
        If(BaseCost_GrandTotal_ChoicePadded_Reverb, 0D),
        If(BaseCost_GrandTotal_Leather_Reverb, 0D),
        If(BaseCost_GrandTotal_LeatherPadded_Reverb, 0D),
        If(BaseCost_GrandTotal_Choice_eBay, 0D),
        If(BaseCost_GrandTotal_ChoicePadded_eBay, 0D),
        If(BaseCost_GrandTotal_Leather_eBay, 0D),
        If(BaseCost_GrandTotal_LeatherPadded_eBay, 0D),
        If(BaseCost_GrandTotal_Choice_Etsy, 0D),
        If(BaseCost_GrandTotal_ChoicePadded_Etsy, 0D),
        If(BaseCost_GrandTotal_Leather_Etsy, 0D),
        If(BaseCost_GrandTotal_LeatherPadded_Etsy, 0D),
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
        RetailPrice_LeatherPadded_Etsy,
        notes
    )
    End Sub
End Module
