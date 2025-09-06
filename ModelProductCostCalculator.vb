Public Class ModelProductCostCalculator
    Public Shared Function CalculateMaterialNeededSqIn(model As Model) As Decimal
        Dim d = model.Depth + 1.25D
        Dim h = model.Height + 1.25D
        Dim w = model.Width + 1.25D
        Return (d * h * 2D) + ((d + h + h + h) * w)
    End Function

    Public Shared Function CalculateTotalCost(
        model As Model,
        materialCostPerSqIn As Decimal,
        laborRate As Decimal,
        productionTimeHours As Decimal,
        packagingCost As Decimal,
        shippingCost As Decimal,
        marketplaceFeePercent As Decimal,
        profitMarginPercent As Decimal
    ) As Decimal
        Dim materialNeeded = CalculateMaterialNeededSqIn(model)
        Dim materialCost = materialCostPerSqIn * materialNeeded
        Dim designFeaturesCost = model.DesignFeatures.Sum(Function(f) f.AddedPrice)
        'Dim paddingCost = If(model.HasPadding, 0D, 0D)  ' Add padding logic if needed
        Dim laborCost = laborRate * productionTimeHours
        Dim subtotal = materialCost + designFeaturesCost + laborCost + packagingCost + shippingCost '+ paddingCost
        Dim profit = subtotal * profitMarginPercent / 100D
        Dim marketplaceFee = (subtotal + profit) * marketplaceFeePercent / 100D
        Return subtotal + profit + marketplaceFee
    End Function
End Class
