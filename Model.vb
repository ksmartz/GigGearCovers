Public Class Model
    Public Property ModelId As Integer
    Public Property SeriesId As Integer
    Public Property ModelName As String ' Unique per Series
    Public Property Width As Decimal
    Public Property Depth As Decimal
    Public Property Height As Decimal

    Public Property TotalFabricSquareInches As Decimal ' <-- Store permanently
    Public Property DesignFeatures As New List(Of ModelDesignFeatures)
    Public Property HasPadding As Boolean
End Class

