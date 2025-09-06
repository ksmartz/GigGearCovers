Public Class BrandDisplayItem
    Public Property PK_FabricBrandNameId As Integer
    Public Property BrandName As String
    Public Property IsSupplierBrand As Boolean

    Public ReadOnly Property DisplayText As String
        Get
            If IsSupplierBrand Then
                Return "★ " & BrandName ' Or use [Supplier] prefix, or color in DrawItem
            Else
                Return BrandName
            End If
        End Get
    End Property
End Class