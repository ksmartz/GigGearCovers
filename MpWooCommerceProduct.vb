Public Class MpWooCommerceProduct
    Public Property name As String ' REQUIRED for parent product (Product title)
    Public Property type As String ' REQUIRED for parent product ("variable" for parent product)
    Public Property status As String ' REQUIRED for parent product ("publish" or "draft")
    Public Property featured As Boolean? ' OPTIONAL for parent product (Mark as featured)
    Public Property catalog_visibility As String ' OPTIONAL for parent product (e.g., "visible", "catalog", "search", "hidden")
    Public Property description As String ' OPTIONAL for parent product (Full product description)
    Public Property short_description As String ' OPTIONAL for parent product (Short description)

    Public Property parent_sku As String
    Public Property sku As String ' OPTIONAL for parent product (Stock Keeping Unit, unique identifier)
    Public Property regular_price As String ' OPTIONAL for parent product (Parent price is not used for variable products; set on variations)
    Public Property sale_price As String ' OPTIONAL for parent product (Set on variations)
    Public Property date_on_sale_from As String ' OPTIONAL for parent product
    Public Property date_on_sale_to As String ' OPTIONAL for parent product
    Public Property manage_stock As Boolean? ' OPTIONAL for parent product (Usually managed at variation level)
    Public Property stock_quantity As Integer? ' OPTIONAL for parent product (Usually managed at variation level)
    Public Property stock_status As String ' OPTIONAL for parent product ("instock", "outofstock", "onbackorder")
    Public Property backorders As String ' OPTIONAL for parent product
    Public Property sold_individually As Boolean? ' OPTIONAL for parent product
    Public Property weight As String ' OPTIONAL for parent product (Usually set on variations)
    Public Property dimensions As MpWooDimensions ' OPTIONAL for parent product (Usually set on variations)
    Public Property shipping_class As String ' OPTIONAL for parent product
    Public Property reviews_allowed As Boolean? ' OPTIONAL for parent product
    Public Property average_rating As String ' OPTIONAL for parent product (Read-only, ignored on create)
    Public Property rating_count As Integer? ' OPTIONAL for parent product (Read-only, ignored on create)
    Public Property related_ids As List(Of Integer) ' OPTIONAL for parent product
    Public Property upsell_ids As List(Of Integer) ' OPTIONAL for parent product
    Public Property cross_sell_ids As List(Of Integer) ' OPTIONAL for parent product
    Public Property parent_id As Integer? ' OPTIONAL for parent product (Used for child/variation products)
    Public Property purchase_note As String ' OPTIONAL for parent product
    Public Property categories As List(Of MpWooCategory) ' OPTIONAL for parent product (Recommended for organization)
    Public Property tags As List(Of MpWooTag) ' OPTIONAL for parent product
    Public Property images As List(Of MpWooImage) ' OPTIONAL for parent product (Recommended for display)
    Public Property attributes As List(Of MpWooAttribute) ' REQUIRED for parent product (Defines variation attributes, e.g., Fabric, Color)
    Public Property default_attributes As List(Of MpWooDefaultAttribute) ' OPTIONAL for parent product (Default selected attributes)
    Public Property variations As List(Of Integer) ' OPTIONAL for parent product (Usually managed by WooCommerce after creation)
    Public Property grouped_products As List(Of Integer) ' OPTIONAL for parent product (For grouped product type)
    Public Property menu_order As Integer? ' OPTIONAL for parent product
    Public Property virtual As Boolean? ' OPTIONAL for parent product
    Public Property downloadable As Boolean? ' OPTIONAL for parent product
    Public Property downloads As List(Of MpWooDownload) ' OPTIONAL for parent product
    Public Property download_limit As Integer? ' OPTIONAL for parent product
    Public Property download_expiry As Integer? ' OPTIONAL for parent product
    Public Property shipping_required As Boolean? ' OPTIONAL for parent product
    Public Property shipping_taxable As Boolean? ' OPTIONAL for parent product
    Public Property shipping_class_id As Integer? ' OPTIONAL for parent product
    Public Property meta_data As List(Of MpWooMetaData) ' OPTIONAL for parent product (Custom fields)


    Public Class MpWooDimensions
        Public Property length As String
        Public Property width As String
        Public Property height As String
    End Class

    Public Class MpWooCategory
        Public Property id As Integer?
        'Public Property name As String
    End Class

    Public Class MpWooTag
        Public Property id As Integer?
        Public Property name As String
    End Class

    Public Class MpWooImage
        Public Property id As Integer?
        Public Property src As String
        Public Property name As String
        Public Property alt As String
        Public Property position As Integer?
    End Class

    Public Class MpWooDefaultAttribute
        Public Property id As Integer? ' OPTIONAL for parent product (Attribute ID, rarely needed)
        Public Property name As String ' REQUIRED if used (Attribute name, e.g., "Color")
        Public Property optionValue As String ' REQUIRED if used (Default value for this attribute, e.g., "Red")
    End Class

    Public Class MpWooDownload
        Public Property id As String
        Public Property name As String
        Public Property file As String
    End Class

    Public Class MpWooMetaData
        Public Property key As String
        Public Property value As String
    End Class

End Class
