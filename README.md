# Agents Commissions VBA macro

In this project a we'll explore a macro that was designed based on a request from accounting in order to report agents' commissions.

Since the data was originated from internal systems, I decided to generate dummy data using python. In order to see the process of generating data, please take a look [here] (https://github.com/mihaiishere/Agents-commissions/blob/main/agents_data.ipynb).

## Excel Planning and Presentation

Since the accountants copy the data from a separate excel file, I decided to create 2 macros (extraction and processing) and 3 sheets. 

In the first sheet I designed the buttons:
![Image](https://github.com/user-attachments/assets/6778f420-e7a7-47a4-a629-62f9d656fe78)

In the second sheet is the data extracted, summarized using a Pivot and on the right I created the header of the report that was requested. Both the Pivot and header will be used later in the script.
![Image](https://github.com/user-attachments/assets/8fffeaa0-1299-4a4a-89da-f38aca15f861)

The third sheet will contain the processed data as requested by accounting. Rows have been hidden for a better overview.
![Image](https://github.com/user-attachments/assets/3f2d577d-6fd6-4591-94e4-c0c94fecbed3)


## The Macro

Now, let's explore what's behind those buttons.


The first VBA module will be the extraction macro.
```
Sub agentsdata()
    'We start by declaring our variables
    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsTaxe As Worksheet
    Dim taxe As Workbook
    Dim lastRow As Long

    Dim rng As Range
    Dim cell As Range
    
    ' In order to be sure that the data is not contaminated we clear columns A to K in the data sheet
    Set wb = ThisWorkbook
    Set wsData = wb.Sheets("data")
    wsData.Range("A:K").Clear
    
    ' Open the excel with data
    Set taxe = Workbooks.Open(wb.Path & "\agent_data.xlsx")
    Set wsTaxe = taxe.Sheets(1)

    ' The next 3 code sections are not needed for our set of data, but the original one had extra types of commissions and all the amounts were negative

    ' Filter the data excel column J to extract only the types of commissions we need
    wsTaxe.Range("A1:K1").AutoFilter Field:=10, Criteria1:=Array("CAS", "CASS", "NET COMMISSION", "TAX"), Operator:=xlFilterValues

    ' Copy filtered data to data sheet column A
    wsTaxe.AutoFilter.Range.Copy wsData.Range("A1")
    taxe.Close False
    
    ' Find and Replace "-" in column K of data sheet
    With wsData.Columns("K")
        .Replace What:="-", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
    End With
    
    
    
     ' Define the range of cells in column E
    Set rng = ThisWorkbook.Sheets("Data").Range("E:E")
    
    ' Loop through each cell in the range and convert it to number in order to not see odd values due to the length
    For Each cell In rng
        If Not IsEmpty(cell) And IsNumeric(cell.Value) Then
            cell.Value = CDbl(cell.Value)
        End If
    Next cell
    
    wsData.Range("E:E").NumberFormat = "0"
    
    ' Refresh Pivot and remove any filter

    wsData.PivotTables("PivotTable1").RefreshTable
    
    MsgBox "Data has been extracted!", vbInformation
    
End Sub
```

Once we have the pop-up confirmation on screen, macro 2 can be pressed.

```
Sub Agentscommissions()
    ' We start by declaring our variables
    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsRezultat As Worksheet
    Dim pvtField As PivotField
    Dim lastRow As Long
    Dim grandTotalRow As Long
    Dim newRow As Long
    
    ' Set workbook and worksheets
    Set wb = ThisWorkbook
    Set wsData = wb.Sheets("data")
    Set wsRezultat = wb.Sheets("result")
    
    ' Delete all data from Rezultat Final
    wsRezultat.Cells.Clear

    ' The below part of the code will also be repeated for the other business line combination, but for a better overview, that part of the code is not repeated here
    ' Filter Pivot active field "Business Line" with "1", "2", and "3"
    Set pvtField = wsData.PivotTables("PivotTable1").PivotFields("Business Line")
    pvtField.ClearAllFilters
    pvtField.PivotItems("1").Visible = True
    pvtField.PivotItems("2").Visible = True
    pvtField.PivotItems("3").Visible = True
    pvtField.PivotItems("(blank)").Visible = False
    
    ' Copy Range U1 into Rezultat final first column with a free row. this and the next copy will set our header
    lastRow = wsRezultat.Cells(wsRezultat.Rows.Count, "B").End(xlUp).Row
    newRow = lastRow + 1
    wsData.Range("U1").Copy wsRezultat.Range("B" & newRow)
    
    ' Copy Range U2:AG2 below it
    wsData.Range("U2:AC2").Copy wsRezultat.Range("B" & newRow + 1)
    
    ' Find "Grand Total" row in the pivot
    grandTotalRow = wsData.Cells.Find(What:="Grand Total", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    grandTotalRow = grandTotalRow - 1
    
    ' Copy filtered data from N5:S to result
    Dim dataRange As Range
    Set dataRange = wsData.Range("N5:S" & grandTotalRow)
    
    ' Copy to Rezultat Final with specified column mapping
    wsRezultat.Cells(newRow + 2, "B").Resize(dataRange.Rows.Count, 1).Value = wsData.Range("N5:N" & grandTotalRow).Value
    wsRezultat.Cells(newRow + 2, "G").Resize(dataRange.Rows.Count, 1).Value = wsData.Range("O5:O" & grandTotalRow).Value
    wsRezultat.Cells(newRow + 2, "H").Resize(dataRange.Rows.Count, 1).Value = wsData.Range("P5:P" & grandTotalRow).Value
    wsRezultat.Cells(newRow + 2, "J").Resize(dataRange.Rows.Count, 1).Value = wsData.Range("Q5:Q" & grandTotalRow).Value
    wsRezultat.Cells(newRow + 2, "I").Resize(dataRange.Rows.Count, 1).Value = wsData.Range("R5:R" & grandTotalRow).Value
    wsRezultat.Cells(newRow + 2, "F").Resize(dataRange.Rows.Count, 1).Value = wsData.Range("S5:S" & grandTotalRow).Value

    ' Add formulas and values to result
    lastRow = wsRezultat.Cells(wsRezultat.Rows.Count, "B").End(xlUp).Row
    With wsRezultat
        .Range("B" & newRow + 1 & ":B" & lastRow).NumberFormat = "0"
        .Range("B" & newRow + 1 & ":B" & lastRow).Borders(xlEdgeLeft).Weight = xlThick
        .Range("J" & newRow + 1 & ":J" & lastRow).Borders(xlEdgeRight).Weight = xlThick
        .Range("E" & newRow + 1 & ":E" & lastRow).NumberFormat = "dd/mm/yyyy"
        .Range("C" & newRow + 1 & ":C" & lastRow).Formula = "=VLOOKUP(B" & newRow + 1 & ", Data!E:G, 2, FALSE)"
        .Range("D" & newRow + 1 & ":D" & lastRow).Formula = "=VLOOKUP(B" & newRow + 1 & ", Data!E:G, 3, FALSE)"
        .Range("E" & newRow + 1 & ":E" & lastRow).Formula = "=VLOOKUP(B" & newRow + 1 & ", Data!E:H, 4, FALSE)"
        wsData.Range("U3:AG3").Copy wsRezultat.Range("B" & lastRow + 1)
        .Range("F" & lastRow + 1).Formula = "=SUM(F" & newRow + 1 & ":F" & lastRow & ")"
        .Range("G" & lastRow + 1).Formula = "=SUM(G" & newRow + 1 & ":G" & lastRow & ")"
        .Range("H" & lastRow + 1).Formula = "=SUM(H" & newRow + 1 & ":H" & lastRow & ")"
        .Range("I" & lastRow + 1).Formula = "=SUM(I" & newRow + 1 & ":I" & lastRow & ")"
        .Range("J" & lastRow + 1).Formula = "=SUM(J" & newRow + 1 & ":J" & lastRow & ")"
        

    End With
    
    wsRezultat.Range("F:N").NumberFormat = "#,##0.00"

    MsgBox "Data processed successfully!", vbInformation
    
End Sub
```

![Image](https://github.com/user-attachments/assets/26d9ad94-14ed-432c-aa5f-9b3ba9b9726d)

If you reached this point, then you will have the data processed as shown at the beginning in sheet result.
