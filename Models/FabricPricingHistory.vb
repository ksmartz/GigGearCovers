Public Class FabricPricingHistory
    Public Property PK_PricingHistoryId As Integer
    Public Property FK_SupplierProductNameDataId As Integer
    Public Property DateFrom As Date
    Public Property DateTo As Nullable(Of Date)
    Public Property ShippingCost As Decimal?
    Public Property CostPerLinearYard As Decimal?
    Public Property CostPerSquareInch As Decimal?
    Public Property WeightPerSquareInch As Decimal?
End Class