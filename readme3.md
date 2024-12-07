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

![Equity_Screening_Tool_Checkboxes.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool_Checkboxes.jpg?raw=true)

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

The *Clear All Filters* checkbox rechecks the Market Cap Size checkboxes, unfilters the data, and sets the Sector, Sub_Industries and Equities to "All". 

    ' Reset all checkboxes and ComboBoxes to initial values and remove all filters
    Private Sub CheckBox_Clear_All_Filters_Change()
        Dim wb1 As Workbook
        Dim ws1 As Worksheet

        Set wb1 = ThisWorkbook
        Set ws1 = wb1.Worksheets("Equities")
        ws1.Activate

        On Error Resume Next

        ' If the Clear All Filters CheckBox is checked
        If ActiveSheet.OLEObjects("CheckBox_Clear_All_Filters").Object.Value = True Then
            ' Set the checkboxes for Market Cap sizes to true
            ActiveSheet.OLEObjects("CheckBox_SmallCap").Object.Value = True
            ActiveSheet.OLEObjects("CheckBox_MidCap").Object.Value = True
            ActiveSheet.OLEObjects("CheckBox_LargeCap").Object.Value = True
        
            ' Set initial values for the Sector, Sub-Industries, and Equities ComboBoxes
            ActiveSheet.ComboBox_Sectors.Value = "All Sectors"
            ActiveSheet.ComboBox_Sub_Industries = "All Sub-Industries"
            ActiveSheet.ComboBox_Equities = "All Equities"
 
            ActiveSheet.ShowAllData
        End If

        ws1.Range("A1").Select
    
    End Sub

### 10. Sort Columns

![Equity_Screening_Tool_Sort.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool_Sort.jpg?raw=true)

We can sort any of the columns using the arrow shapes which calls sorting procedures. One of the Up arrows calls the *Sort_1_Asc* procedure and the Down arrow calls the *Sort_1_Desc* procedure. They call the *Sort_Sheet* procedure to execute the sort based on parameters passed.

    ' Sort by 1st Dynamic ComboBox Ascending
    Sub Sort_1_Asc()
        Dim wb1 As Workbook
        Dim ws1 As Worksheet
        Dim start_row As Integer
        Dim end_row As Integer

        Set wb1 = ThisWorkbook
        Set ws1 = wb1.Worksheets("Equities")
    
        ws1.Activate

        Sort_Sheet 8, "C", 10, 18, "Asc"
    
    End Sub

    ' Sort by 1st Dynamic ComboBox Descending
    Sub Sort_1_Desc()
        Dim wb1 As Workbook
        Dim ws1 As Worksheet
        Dim start_row As Integer
        Dim end_row As Integer

        Set wb1 = ThisWorkbook
        Set ws1 = wb1.Worksheets("Equities")
    
        ws1.Activate

        Sort_Sheet 8, "C", 10, 18, "Desc"
    
    End Sub

    ' Sort the sheet by curr_col column and whether sort_order is ascending or descending
    Sub Sort_Sheet(start_row As Integer, start_col As Variant, curr_col As Integer, end_col As Integer, sort_order As String)
        Dim wb1 As Workbook
        Dim ws1 As Worksheet
        Dim end_row As Integer

        Set wb1 = ThisWorkbook
        Set ws1 = wb1.Worksheets("Equities")
    
        ws1.Activate

        ' Search for blank row to get last row
        end_row = Search_for_value(start_row, ws1.Name, start_col, "") - 1
    
        ' Set up sort key by curr_col column and whether sort_order is ascending or descending
        ws1.Sort.SortFields.Clear
        If sort_order = "Asc" Then
            ws1.Sort.SortFields.Add2 Key:=ws1.Range(ws1.Cells(start_row, curr_col), ws1.Cells(end_row, curr_col)) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        Else
            ws1.Sort.SortFields.Add2 Key:=ws1.Range(ws1.Cells(start_row, curr_col), ws1.Cells(end_row, curr_col)) _
            , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        End If

        With ws1.Sort
            .SetRange ws1.Range(ws1.Cells(start_row, 2), ws1.Cells(end_row, end_col))
            .Header = xlGuess
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    
    End Sub

### 11. Clear Data

![Equity_Screening_Tool_Clear_Screen.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool_Clear_Screen.jpg?raw=true)

The eraser icon shape clears the sheet using the *Clear_Sheet* procedure.

    ' Clear the whole sheet
    Sub Clear_Sheet()
        Dim wb1 As Workbook
        Dim ws1 As Worksheet
        Dim start_row As Integer
        Dim end_row As Integer
        Dim start_col As Integer
        Dim end_col As Integer

        Set wb1 = ThisWorkbook
        Set ws1 = wb1.Worksheets("Equities")
        ws1.Activate
    
        On Error Resume Next
        ws1.ShowAllData
    
        ThisWorkbook.DisableApplication
    
        start_row = 8
        start_col = 3
        end_col = 17
    
        ' Look for first row that contains a blank in column start_col
        end_row = Search_for_value(start_row, ws1.Name, start_col, "") - 1
    
        ' Clear the sheet
        Clear_Section start_row, ws1.Name, end_row, start_col, end_col
    
        ThisWorkbook.EnableApplication
    
    End Sub

The screen is cleared by clicking the button.

![Equity_Screening_Tool_Query_Results_Clear.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool_Query_Results_Clear.jpg?raw=true)

### 12. Show / Hide Description Columns

![Equity_Screening_Tool_Query_Results2_Hide.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool_Query_Results2_Hide.jpg?raw=true)

In order to see all the data at once, we can hide columns "E" to "G". If the *Hide Desc Cols* Option Button is clicked, it calles the *Hide_Desc_Cols* procedure. The *Show Desc Cols* Option Button is clicked, it calls the *Show_Desc_Cols* procedure.

    ' If Hide Desc Cols Option button is clicked, hide the description columns E to G
    Private Sub OptionButton_Hide_Desc_Cols_Click()
        Hide_Desc_Cols
    End Sub

    ' Hide Description Columns E to G
    Sub Hide_Desc_Cols()
        Dim wb1 As Workbook
        Dim ws1 As Worksheet

        Set wb1 = ThisWorkbook
        Set ws1 = wb1.Worksheets("Equities")
        ws1.Activate
    
        If ActiveSheet.Columns("G").Hidden = False Then
            ActiveSheet.Columns("E:G").Hidden = True
            ActiveSheet.Shapes("Sort_Asc_Sector").Visible = False
            ActiveSheet.Shapes("Sort_Desc_Sector").Visible = False
            ActiveSheet.Shapes("Sort_Asc_Sub_Ind").Visible = False
            ActiveSheet.Shapes("Sort_Desc_Sub_Ind").Visible = False
            ActiveSheet.Shapes("Sort_Asc_Comp_Name").Visible = False
            ActiveSheet.Shapes("Sort_Desc_Comp_Name").Visible = False
            ActiveSheet.Shapes("Sort_Asc_Market_Cap").Visible = False
            ActiveSheet.Shapes("Sort_Desc_Market_Cap").Visible = False
        End If
    
        ws1.Range("A1").Select

    End Sub

    ' If Show Desc Cols Option button is clicked, show the description columns E to G
    Private Sub OptionButton_Show_Desc_Cols_Click()
        Show_Desc_Cols
    End Sub

    ' Show Description Columns E to G
    Sub Show_Desc_Cols()
        Dim wb1 As Workbook
        Dim ws1 As Worksheet

        Set wb1 = ThisWorkbook
        Set ws1 = wb1.Worksheets("Equities")
        ws1.Activate
    
        If ActiveSheet.Columns("G").Hidden = True Then
            ActiveSheet.Columns("E:G").Hidden = False
            ActiveSheet.Shapes("Sort_Asc_Sector").Visible = True
            ActiveSheet.Shapes("Sort_Desc_Sector").Visible = True
            ActiveSheet.Shapes("Sort_Asc_Sub_Ind").Visible = True
            ActiveSheet.Shapes("Sort_Desc_Sub_Ind").Visible = True
            ActiveSheet.Shapes("Sort_Asc_Comp_Name").Visible = True
            ActiveSheet.Shapes("Sort_Desc_Comp_Name").Visible = True
            ActiveSheet.Shapes("Sort_Asc_Market_Cap").Visible = True
            ActiveSheet.Shapes("Sort_Desc_Market_Cap").Visible = True
        End If
    
        ws1.Range("A1").Select
    
    End Sub

The Description Columns were hidden by clicking the *Hide Desc Cols* Option Button.

![Equity_Screening_Tool_Query_Results2_Hide.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool_Query_Results2_Hide.jpg?raw=true)
    
