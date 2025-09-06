Public Class MpWooAttribute
    Public Property name As String ' REQUIRED for parent product attribute (e.g., "Fabric", "Color")
    Public Property variation As Boolean ' REQUIRED for parent product attribute (True if used for variations)
    Public Property visible As Boolean ' REQUIRED for parent product attribute (True to show on product page)
    Public Property options As List(Of String) ' REQUIRED for parent product attribute (List of possible values, e.g., {"Red", "Blue"})
End Class
