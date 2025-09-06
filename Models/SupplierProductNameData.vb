Imports System.ComponentModel.DataAnnotations

Public Class SupplierProductNameData

    Public Property PK_SupplierProductNameDataId As Integer
    Public Property FK_SupplierNameId As Integer
    Public Property FK_FabricBrandProductNameId As Integer
    Public Property FK_FabricTypeNameId As Integer
    Public Property WeightPerLinearYard As Decimal?
    Public Property SquareInchesPerLinearYard As Decimal?
    Public Property FabricRollWidth As Decimal?
    Public Property TotalYards As Decimal?
    Public Property IsActiveForMarketplace As Boolean
End Class