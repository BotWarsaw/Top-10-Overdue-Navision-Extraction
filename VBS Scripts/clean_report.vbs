Function CheckIfSheetExists(SheetName As String) As Boolean
      CheckIfSheetExists = False
      For Each WS In Worksheets
        If SheetName = WS.Name Then
          CheckIfSheetExists = True
          Exit Function
        End If
      Next WS
End Function

Function Clean_Report(filePath As String)

Application.ScreenUpdating = false
Application.DisplayAlerts = False

Dim wb As Workbook
Set wb = Application.Workbooks.Open(filePath)

Dim report As Worksheet
Dim info As Worksheet


If CheckIfSheetExists("Report") = True Then
    wb.Sheets("Report").Delete
    wb.Sheets.Add.Name = "Report"
Else

    wb.Sheets.Add.Name = "Report"
End If


Set report = wb.Sheets("Report")
Set sh_data = wb.Sheets("Sheet1")


for i = 2 to 20

if sh_data.cells(i,1).value = "No." then
lrow=i
end if
next i


customer_clm = 1
for i = 1 to 30
if sh_data.cells(lrow,i).value = "Not Due" then
not_due=i
end if



if sh_data.cells(lrow,i).value = "0-30 days" then
first_range =i
end if

if sh_data.cells(lrow,i).value = "31-60 days" then
second_range =i
end if

if sh_data.cells(lrow,i).value = "61-90 days" then
third_range =i
end if

if sh_data.cells(lrow,i).value = "91-120 days" then
fourth_range =i
end if

if sh_data.cells(lrow,i).value = "121-180 days" then
fifth_range =i
end if


if sh_data.cells(lrow,i).value = "181-365 days" then

sixth_range =i

end if



if sh_data.cells(lrow,i).value = "Over 365 days" then
seventh_range =i
end if

next i


report.Cells(1, 1).Value = "Contractor"
report.Cells(1, 2).Value = "Current"
report.Cells(1, 3).Value = "0-30"
report.Cells(1, 4).Value = "31-60"
report.Cells(1, 5).Value = "61-90"
report.Cells(1, 6).Value = "91-120"
report.Cells(1, 7).Value = "121-180"
report.Cells(1, 8).Value = "181-365"
report.Cells(1, 9).Value = "Over 365"

last_row = sh_data.Cells(Rows.Count, 1).End(xlUp).Row - 2
a = 2
For i = lrow + 2 To last_row

report.Cells(a, 1).Value = sh_data.Cells(i, 2).Value
report.Cells(a, 2).Value = sh_data.Cells(i, not_due).Value
report.Cells(a, 3).Value = sh_data.Cells(i, first_range).Value
report.Cells(a, 4).Value = sh_data.Cells(i, second_range).Value
report.Cells(a, 5).Value = sh_data.Cells(i, third_range).Value
report.Cells(a, 6).Value = sh_data.Cells(i, fourth_range).Value
report.Cells(a, 7).Value = sh_data.Cells(i, fifth_range).Value
report.Cells(a, 8).Value = sh_data.Cells(i, sixth_range).Value
report.Cells(a, 9).Value = sh_data.Cells(i, seventh_range).Value
a = a + 1
Next i


Application.ScreenUpdating = true
Application.DisplayAlerts = true




End Function