Attribute VB_Name = "clear_wb_formatting"
Sub clear_wb_formatting()

    '   A simple macro that loops through each sheet in the given workbook clearing cell formatting (excluding number formats).

    '   If you need everything reset "wSheet.Cells.ClearFormats" (see below) will clear everything.
    '   Conditional formatting can also be removed using the below, this is commented out by default

    '   Updates the default "Normal" formatting for the whole workbook
    With ActiveWorkbook.Styles("Normal").Font

        .Name = "Calibri"   'this can be changed to any font required
        .Size = 11          'as can the size
        .Bold = False
        .Italic = False
        .Underline = xlUnderlineStyleNone
        .Strikethrough = False
        .Superscript = False
        .ColorIndex = xlAutomatic

    End With

    '   Loops through each sheet and clears the existing formatting,
    '   Returns formatting to the 'new' default set above.

    Dim wSheet As Worksheet 'variable used in the 'For Loop' below

    For Each wSheet In ActiveWorkbook.Worksheets

        '       If all you need is to clear everything uncomment the below - this will however reset date formats to numbers
        '       wSheet.Cells.ClearFormats

        '       Colour formatting
        wSheet.Cells.Interior.ColorIndex = xlNone
        wSheet.Cells.Font.ColorIndex = xlAutomatic

        '       Border formatting
        wSheet.Cells.Borders.LineStyle = xlNone
        wSheet.Cells.Borders.ColorIndex = xlNone

        '       Hyperlinks and underlines
        wSheet.Cells.ClearHyperlinks
        wSheet.Cells.Font.Underline = False

        '       Cell alignment
        '       wSheet.Cells.VerticalAlignment = xlTop 'reset alignment
        '       wSheet.Cells.HorizontalAlignment = xlLeft 'reset alignment

        '       Removes all conditional formatting
        '       wSheet.Cells.FormatConditions.Delete 'clear all conditional formating

    Next wSheet

End Sub
