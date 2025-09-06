Public Class ModelCostUpdateResults
    Public Property ModelId As Integer
    Public Property ModelName As String
    Public Property SeriesName As String
    Public Property ManufacturerName As String
    Public Property Updated As Boolean
    Public Property Message As String

    Public Sub New(modelId As Integer, modelName As String, seriesName As String, manufacturerName As String, updated As Boolean, message As String)
        Me.ModelId = modelId
        Me.ModelName = modelName
        Me.SeriesName = seriesName
        Me.ManufacturerName = manufacturerName
        Me.Updated = updated
        Me.Message = message
    End Sub
End Class
