Imports System.Reflection

Public Enum EquipmentType
    GuitarAmplifier
    MusicKeyboard
End Enum

Public Class ModelSeries
    Public Property SeriesId As Integer
    Public Property ManufacturerId As Integer
    Public Property SeriesName As String ' Unique per Manufacturer
    Public Property EquipmentTypeId As Integer ' FK to EquipmentType table
    Public Property ModelList As New List(Of Model)
End Class
