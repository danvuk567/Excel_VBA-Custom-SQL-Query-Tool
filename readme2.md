# Excel VBA Custom SQL Query Tool (Part Two)

![Equity_Screening_Tool.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool.jpg?raw=true)

### 4. Ticker List

If the checkbox called *CheckBox_Tickers* is checked and the certain Tickers are typed into the "B" column, they are used to filter the Equity data to only return the data for the Tickers listed. The data is filtered by the Ticker list that is appended to the strSQL1 SQL statement within the *Get_Equity_Data* procedure. Once the query is executed, it is sorted to match the order in the "B" column by the sort order in the "R" column which is derived by using the **MATCH** Excel function. The sort order in the "R" column is visually hidden by using a white text font.

![Equity_Screening_Tool_Ticker_List.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool_Ticker_List.jpg?raw=true)

        ' Use Ticker List in B columnn to query if the Tickers CheckBox is checked
        If ActiveSheet.CheckBox_Tickers.Value = True Then
            ' Check for first row that contains a blank in column B
            end_row = Search_for_value(start_row, ws1.Name, "B", "") - 1

            ' Check if specific Tickers were specified by checking if end_row > start_row
            If end_row > start_row Then
                ticker_list = ""

                ' Build the Ticker list from the Tickers in B column
                For i = start_row To end_row
                    If i = end_row Then
                        ticker_list = ticker_list & "'" & ws1.Range("B" & i).Value & "'"
                    Else
                        ticker_list = ticker_list & "'" & ws1.Range("B" & i).Value & "',"
                    End If
                    ' Store the position of Tickers in column D for later sorting purposes to match the order for column B
                    ws1.Range("R" & i).Formula = "=MATCH($D$" & i & ",$B$" & start_row & ":$B$" & end_row & ", FALSE)"
                    ' Hide the position by whiting the text font
                    ws1.Range("R" & i).Font.ThemeColor = xlThemeColorDark1
                Next i

            Else
                ' Set ticker_list to "XXX" to pull nothing if the list is empty while Tickers CheckBox is checked
                ticker_list = "'XXX'"
            End If

            ' Include Ticker list as a condition in strSQL1
            strSQL1 = " AND Ticker_Name IN (" & ticker_list & ")"
        End If
       
        strSQL = strSQL & strSQL1 & " ORDER BY 1,2"

### 5. Clear Tickers

If the *Clear Tickers* shape is clicked, it calls the procedure *Clear_Tickers*. The procedure clears all the Tickers in the "B" column and uses the custom procedures *Search_for_value* and *Clear_Section* that have been defined in .

        ' Clear the Tickers in the B column
        Sub Clear_Tickers()
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
            start_col = 2
            end_col = 2
    
            ' Look for first row that contains a blank in column start_col
            end_row = Search_for_value(start_row, ws1.Name, start_col, "") - 1
    
            ' Clear column B
            Clear_Section start_row, ws1.Name, end_row, start_col, end_col
    
            ThisWorkbook.EnableApplication

        End Sub



