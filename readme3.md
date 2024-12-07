# Excel VBA Custom SQL Query Tool (Part Three)

![Equity_Screening_Tool.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool.jpg?raw=true)

### 7. Stop Query

The stop query shape calls the procedure *Cancel_Running_Query* which simply interrupts the query from continuing the run.

![Equity_Screening_Tool_Stop_Query.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool_Stop_Query.jpg?raw=true)

    ' Interrupt the query with a keystroke break
    Sub Cancel_Running_Query()
        Dim wb1 As Workbook
        Dim ws1 As Worksheet

        Set wb1 = ThisWorkbook
        Set ws1 = wb1.Worksheets("Equities")
        ws1.Activate

        Application.SendKeys "^{BREAK}"
    
        ws1.Range("A1").Select
    
    End Sub

### 8. Query Results

The current query returns the results for the *Information Technology* Sector and the *Semiconductors* Sub-Industry for December 6th 2024.

![Equity_Screening_Tool_Query_Results.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool_Query_Results.jpg?raw=true)

The current query returns the results for the Tickers in column "B" for December 6th 2024.

![Equity_Screening_Tool_Query_Results.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool_Query_Results2.jpg?raw=true)

### 9. Market Cap Size and Clear Filters

The 3 Market Cap Size checkboxes will filter the data by the "H" column using the procedure *Filter_Cap*.

    ' This procedure will filter the sheet by Market Cap Size if any of the Market Cap checkboxes are selected
    Sub Filter_Cap()
        Dim wb1 As Workbook
        Dim ws1 As Worksheet
    Dim i As Integer
    Dim start_row As Integer
    Dim end_row As Integer
    Dim filter_col As Integer
    Dim Cap_Array() As String
    Dim rng As Range
    Dim start_col As Integer
    Dim end_col As Integer

    Set wb1 = ThisWorkbook
    Set ws1 = wb1.Worksheets("Equities")
    ws1.Activate

    start_row = 8
    start_col = 3
    end_col = 17

    ' Filter column index for the Market Cap Size column
    filter_col = 7
    
    ' Find the first blank row to filter the data up until that row
    end_row = Search_for_value(start_row, ws1.Name, start_col, "") - 1
    
    ' Initialize the Market Cap Size array dynamically for the filter values
    i = 0
    ReDim Cap_Array(0 To 2)
    Cap_Array(0) = ""
 
    ' Check if Small Cap checkbox is selected and populate the array accordingly
    If ActiveSheet.OLEObjects("CheckBox_SmallCap").Object.Value = True Then
        Cap_Array(i) = "Small Cap"
        i = i + 1
    End If
    
    ' Check if Mid Cap checkbox is selected and populate the array accordingly
    If ActiveSheet.OLEObjects("CheckBox_MidCap").Object.Value = True Then
        Cap_Array(i) = "Mid Cap"
        i = i + 1
    End If
    
    ' Check if Large Cap checkbox is selected and populate the array accordingly
    If ActiveSheet.OLEObjects("CheckBox_LargeCap").Object.Value = True Then
        Cap_Array(i) = "Large Cap"
        i = i + 1
    End If
    
    ' If any checkboxes are selected, adjust the size of the array
    If i > 0 Then
        ReDim Preserve Cap_Array(0 To i - 1) ' Resize to only selected options
        ws1.OLEObjects("CheckBox_Clear_All_Filters").Object.Value = False ' Clear filter indicator
    Else
        ' If no checkboxes are selected, exit the sub (no filtering needed)
        Exit Sub
    End If
   
    ' Define the range to apply the filter
    Set rng = ActiveSheet.Range(ws1.Cells(start_row, start_col), ws1.Cells(end_col, end_row))
    
    ' Apply the filter
    rng.AutoFilter Field:=filter_col, Criteria1:=Cap_Array, Operator:=xlFilterValues

    ws1.Range("A1").Select
    
End Sub

' Apply any filter changes if Small Cap checkbox is checked
Private Sub CheckBox_SmallCap_Change()
    Filter_Cap
End Sub

' Apply any filter changes if Mid Cap checkbox is checked
Private Sub CheckBox_MidCap_Change()
    Filter_Cap
End Sub

' Apply any filter changes if Large Cap checkbox is checked
Private Sub CheckBox_LargeCap_Change()
    Filter_Cap
End Sub
