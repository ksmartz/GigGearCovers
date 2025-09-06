Public Class ModelManufacturers
    Public Property ManufacturerId As Integer
    Public Property Name As String ' Unique
    Public Property SeriesList As New List(Of ModelSeries)
End Class
