Imports System.IO
Imports System.Text

Module WooCommerceCSVGenerator

    ' Escapes a value for CSV (handles commas, quotes, newlines)
    Private Function CsvEscape(value As String) As String
        If value.Contains(",") OrElse value.Contains("""") OrElse value.Contains(vbCr) OrElse value.Contains(vbLf) Then
            value = value.Replace("""", """""")
            Return $"""{value}"""
        End If
        Return value
    End Function

    Public Sub CreateWooCommerceVariationsCsv(filePath As String)
        Dim fabrics = New String() {"Waterproof", "Premium Synthetic Leather"}
        Dim colors = New String() {"Black", "Red", "Blue"}

        Dim extraColumns = New String() {
            "Extra Padding", "2-in-1 Zipper Pocket",
            "Zipperized Amp Handle Cover - Top",
            "Zipperized Amp Handle Cover - Side",
            "Shipping Method", "Priority Queue"
        }

        Dim sb As New StringBuilder()

        ' Header
        Dim header = New List(Of String) From {
            "Type", "Name", "Attribute 1 name", "Attribute 1 value(s)",
            "Attribute 2 name", "Attribute 2 value(s)"
        }
        header.AddRange(extraColumns)
        sb.AppendLine(String.Join(",", header.Select(Function(h) CsvEscape(h))))

        ' Parent product row
        Dim parentRow = New List(Of String) From {
            "variable", "Custom Amp Cover", "Fabric", String.Join("|", fabrics),
            "Color", String.Join("|", colors)
        }
        parentRow.AddRange(extraColumns.Select(Function(c) ""))
        sb.AppendLine(String.Join(",", parentRow.Select(Function(v) CsvEscape(v))))

        ' Variation rows
        For Each fabric In fabrics
            For Each color In colors
                Dim row = New List(Of String) From {
                    "variation", "Custom Amp Cover", "Fabric", fabric, "Color", color,
                    "No", "No", "No", "No", "UPS", "No"
                }
                sb.AppendLine(String.Join(",", row.Select(Function(v) CsvEscape(v))))
            Next
        Next

        File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8)
    End Sub

End Module
