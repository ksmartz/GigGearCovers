Imports System

Module TESTMODULE
    Sub Main()
        'Dim template As String = "{{BASERETAILPRICE_CHOICE + POCKET + ZIPPERHANDLE}}"


        Dim template As String = "| {{BASERETAILPRICE_CHOICE + POCKET}}</li><li> <b><i>Yes</i></b> 2-in-1 Cargo 'Pick Pocket', <b><i>Yes</i></b> Zipper Handle Cover | {{BASERETAILPRICE_CHOICE + POCKET + ZIPPERHANDLE}}</li></ul>"

        Dim desc = ListingHelpers.BuildListingDescription(template, 7, "Reverb", 3)
        Console.WriteLine(desc)
    End Sub


    'Sub Main()
    '    ' Paste your full HTML template here
    '    Dim htmlTemplate As String = "<p>Your <b>🎸{MANUFACTURER_NAME} {SERIES_NAME} {MODEL_NAME}🎸</b> ... (rest of your HTML) ... </p>"
    '    ' Use equipmentTypeId = 1 or 3 to match your table data
    '    Dim description As String = ListingHelpers.BuildListingDescription(htmlTemplate, 7, "Reverb", 3)
    '    Console.WriteLine(description)
    'End Sub
End Module
