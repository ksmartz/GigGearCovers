Imports System.ComponentModel.DataAnnotations

Public Class FabricBrandProductName

    Public Property PK_FabricBrandProductNameId As Integer
    Public Property BrandProductName As String
    Public Property FK_FabricBrandNameId As Integer
    Public Property WeightPerLinearYard As Decimal
    Public Property FabricRollWidth As Decimal
    Public Overridable Property FabricBrandNames As FabricBrandName
    Public Overridable Property SupplierProductNameDatas As ICollection(Of SupplierProductNameData)
End Class