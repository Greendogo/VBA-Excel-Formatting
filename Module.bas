Attribute VB_Name = "Module1"
Sub FormatRowsWithBoldTitles()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim outputCell As Range
    
    ' Set the worksheet where the data is
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to the name of your sheet
    
    ' Store the section titles in an array
    Dim sectionTitles As Variant
    sectionTitles = Array(ws.Cells(1, 1).Value, ws.Cells(1, 2).Value, ws.Cells(1, 3).Value, ws.Cells(1, 4).Value)
    
    ' Find the last row with data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Loop through each row starting from the second row (assuming the first row contains headers)
    For i = 2 To lastRow
        ' Get the values from the corresponding columns
        Dim content As Variant
        content = Array(ws.Cells(i, 1).Value, ws.Cells(i, 2).Value, ws.Cells(i, 3).Value, ws.Cells(i, 4).Value)
        
        ' Build the full text
        Dim fullText As String
        fullText = BuildFullText(sectionTitles, content)
        
        ' Set the output cell where you want the formatted data to appear
        Set outputCell = ws.Cells(i, 5) ' Change column 5 to where you want the output
        
        ' Clear any previous content in the output cell
        outputCell.Clear
        
        ' Insert the full text
        outputCell.Value = fullText
        
        ' Apply bold formatting
        ApplyBoldFormatting outputCell, sectionTitles, content
    Next i
End Sub

' Function to build the full text
Function BuildFullText(sectionTitles As Variant, content As Variant) As String
    Dim fullText As String
    Dim j As Integer
    
    ' Loop through each section to build the text
    fullText = ""
    For j = 0 To UBound(sectionTitles)
        fullText = fullText & sectionTitles(j) & vbNewLine & content(j) & vbNewLine & vbNewLine
    Next j
    
    ' Remove the last two vbNewLine characters
    fullText = Left(fullText, Len(fullText) - Len(vbNewLine & vbNewLine))
    
    ' Return the result
    BuildFullText = fullText
End Function

' Function to apply bold formatting to section titles
Sub ApplyBoldFormatting(outputCell As Range, sectionTitles As Variant, content As Variant)
    Dim startPos As Long
    Dim lastLen As Long
    Dim j As Integer
    Dim newline As Integer
    newline = 2
    
    ' Initialize the starting position
    startPos = 1
    
    ' Loop through each section title and apply bold formatting
    For j = 0 To UBound(sectionTitles)
        lastLen = Len(sectionTitles(j)) + newline
        With outputCell.Characters(startPos, lastLen).Font
            .Bold = True
        End With
        startPos = startPos + lastLen + Len(content(j)) + newline + newline
    Next j
End Sub

