The following code was provided by the instructor of our class:
    ws.Cells(Rows.Count, 1).End(xlUp).Row
    Dim ws As Worksheet
    For Each ws In Worksheets
    Next ws
    For i = 2 To LastRow
    Next i

Everything else was crafted based on code learned in class or from one of the following two sources:
	https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofit
	https://learn.microsoft.com/en-us/office/vba/api/excel.range.numberformat
for the code surrounding autofitting columns and formating a number as a percent, respectively.
This was not copied from these sites and should not violate any plagiarism rules, but I'm including it for complete transparency.