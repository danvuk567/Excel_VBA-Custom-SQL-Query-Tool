# Excel VBA Custom SQL Query Tool (Part Two)

![Equity_Screening_Tool.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool.jpg?raw=true)

### 4. Ticker List

If the checkbox called *CheckBox_Tickers* is checked and the certain Tickers are typed into the "B" column, they are used to filter the Equity data to only return the data for the Tickers listed. The data is filtered by the Ticker list that is appended to the strSQL1 SQL statement within the *Get_Equity_Data* procedure. Once the query is executed, it will be sorted to match the order in the "B" column by the sort order in the "R" column which is derived by using the **MATCH** Excel function. The sort order in the "R" column is visually hidden by using a white text font.

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

![Equity_Screening_Tool_Clear_Ticker_List.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool_Clear_Ticker_List.jpg?raw=true)

If the *Clear Tickers* shape is clicked, it calls the procedure *Clear_Tickers*. The procedure clears all the Tickers in the "B" column and uses the custom procedures *Search_for_value* and *Clear_Section* that have been defined in [Excel VBA Useful Custom Functions](https://github.com/danvuk567/Excel_VBA-Useful-Custom-Functions).

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

### 6. Execute the Query

![Equity_Screening_Tool_Ticker_List.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool_Ticker_List.jpg?raw=true)

The shape with the Magnify (search) icon calls the procedure *Get_Equity_Data* procedure which filters the SQL statement strSQL based on the Sector, Sub-Industry, Equity ComboBoxes, and Ticker list in column "B". It then calls the procedure *Exec_Equity_Data start_row*. 
If there are Tickers listed in the "B" column, it will sort them using the *Sort_Tickers* procedure. The data will populate columns "C" to "H" and any data related to what is selected in the dynamic ComboBox in "J6" to "Q6".

        ' Set up the query and call Exec_Equity_Data procedure
        Sub Get_Equity_Data()
            Dim strSQL As String
            Dim strSQL1 As String
            Dim query_cell As String
            Dim wb1 As Workbook
            Dim ws1 As Worksheet
            Dim ws2 As Worksheet
            Dim i As Integer
            Dim start_row As Integer
            Dim end_row As Integer
            Dim start_col As Variant
            Dim end_col As Variant
            Dim ticker_list As String

            Set wb1 = ThisWorkbook
            Set ws1 = wb1.Worksheets("Equities")
            Set ws2 = wb1.Worksheets("Queries")
    
            ws1.Activate
    
            On Error Resume Next
            ws1.ShowAllData
    
            ThisWorkbook.DisableApplication
    
            start_row = 8
            start_col = "C"
            end_col = "R"
            query_cell = "A5"
    
            ' Look for first row that contains a blank in column start_col
            end_row = Search_for_value(start_row, ws1.Name, start_col, "") - 1
    
            ' Clear the sheet in order to re-populate with query data
            Clear_Section start_row, ws1.Name, end_row, start_col, end_col

            strSQL = ws2.Range(query_cell).Value & "'" & ActiveSheet.ComboBox_Start_Date.Value & "' AND '" & ActiveSheet.ComboBox_End_Date.Value & "' "

            If ActiveSheet.ComboBox_Start_Date.Value <= ActiveSheet.ComboBox_End_Date.Value Then

                ' Add on conditions to strSQL1 depending on what is selected in the Sub-Industries and Equity ComboBox values
                If ActiveSheet.ComboBox_Sectors.Value = "All Sectors" Then
                    strSQL1 = ""
                Else
                    If ActiveSheet.ComboBox_Sub_Industries.Value = "All Sub-Industries" Then
                        strSQL1 = " AND Sector = '" & ActiveSheet.ComboBox_Sectors.Value & "'"
                    Else
                        If ActiveSheet.ComboBox_Equities.Value = "All Equities" Then
                            strSQL1 = " AND Sub_Industry = '" & ActiveSheet.ComboBox_Sub_Industries.Value & "'"
                        Else
                            strSQL1 = " AND Ticker_Name = TRIM(LEFT('" & ActiveSheet.ComboBox_Equities.Value & "',CHARINDEX(':','" & ActiveSheet.ComboBox_Equities.Value & "')-1))"
                        End If
                    End If
                End If


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
        
                ' Call Exec_Equity_Data procedure to query with strSQL and populate data starting at row start_row
                Exec_Equity_Data start_row, strSQL
        
            End If
    
            ' Scroll back to row 8 and column 5 in Freeze Pane
            ActiveWindow.ScrollRow = 8
            ActiveWindow.ScrollColumn = 5

            ThisWorkbook.EnableApplication
    
            ' Sort the Tickers, if any, in B column to match column D
            If (ActiveSheet.CheckBox_Tickers.Value) And (end_row > start_row) Then
                Sort_Tickers start_row, end_row
            End If
    
            ws1.Range("A1").Select
    
        End Sub

        ' Execute the strSQL query and populate the rows starting from the start_row row
        Sub Exec_Equity_Data(start_row As Integer, strSQL As String)
            Dim objMyConn As ADODB.Connection
            Dim objMyRecordset As ADODB.Recordset
            Dim stConn As String
            Dim wb1 As Workbook
            Dim ws1 As Worksheet
            Dim i As Integer
            Dim j As Integer

            Set wb1 = ThisWorkbook
            Set ws1 = wb1.Worksheets("Equities")
            Set ws2 = wb1.Worksheets("Queries")
            ws1.Activate
    
            On Error GoTo ErrorHandler
    
            ' Initialize connection objects
            Set objMyConn = New ADODB.Connection
            Set objMyRecordset = New ADODB.Recordset
            stConn = Set_Conn
            objMyConn.ConnectionString = stConn
  
            ' Query only if Start Date <= End Date
            If ActiveSheet.ComboBox_Start_Date.Value <= ActiveSheet.ComboBox_End_Date.Value Then
    
                ' Get RecordSet for SQL Execution of strSQL
                With objMyConn
                    .CursorLocation = adUseClient 'Necessary for creating disconnected recordset.
                    .Open stConn 'Open connection.
                    Set objMyRecordset = .Execute(strSQL)
                End With

                i = start_row
            
                ' Move through RecordSet
                objMyRecordset.MoveFirst
                Do While Not objMyRecordset.EOF
                    ' Populate the static column values from the RecordSet
                    ws1.Range(ws1.Cells(i, 3), ws1.Cells(i, 3)) = objMyRecordset.Fields("Date")
                    ws1.Range(ws1.Cells(i, 3), ws1.Cells(i, 3)).NumberFormat = "yyyy-mm-dd"
                    ws1.Range(ws1.Cells(i, 4), ws1.Cells(i, 4)) = objMyRecordset.Fields("Ticker_Name")
                    ws1.Range(ws1.Cells(i, 5), ws1.Cells(i, 5)) = objMyRecordset.Fields("Sector")
                    ws1.Range(ws1.Cells(i, 6), ws1.Cells(i, 6)) = objMyRecordset.Fields("Sub_Industry")
                    ws1.Range(ws1.Cells(i, 7), ws1.Cells(i, 7)) = objMyRecordset.Fields("Company_Name")
                    ws1.Range(ws1.Cells(i, 8), ws1.Cells(i, 8)) = objMyRecordset.Fields("Market_Cap")
                    ws1.Range(ws1.Cells(i, 8), ws1.Cells(i, 8)).NumberFormat = "#,##0"
                    ws1.Range(ws1.Cells(i, 9), ws1.Cells(i, 9)) = objMyRecordset.Fields("Market_Cap_Size")
            
                    ' Populate the dynamic column values from the RecordSet which are chosen from the ComboBoxes
                    For j = 1 To 8
                             Select Case ws1.OLEObjects("ComboBox_Score_Fields" & CStr(j)).Object.Value
                                Case "Chg % 1M"
                                    ws1.Range(ws1.Cells(i, (j + 9)), ws1.Cells(i, (j + 9))) = objMyRecordset.Fields("Return_1Mth")
                                    ws1.Range(ws1.Cells(i, (j + 9)), ws1.Cells(i, (j + 9))).NumberFormat = "0.00%"
                                Case "Chg % 3M"
                                    ws1.Range(ws1.Cells(i, (j + 9)), ws1.Cells(i, (j + 9))) = objMyRecordset.Fields("Return_3Mth")
                                    ws1.Range(ws1.Cells(i, (j + 9)), ws1.Cells(i, (j + 9))).NumberFormat = "0.00%"
                                Case "Chg % 6M"
                                    ws1.Range(ws1.Cells(i, (j + 9)), ws1.Cells(i, (j + 9))) = objMyRecordset.Fields("Return_6Mth")
                                    ws1.Range(ws1.Cells(i, (j + 9)), ws1.Cells(i, (j + 9))).NumberFormat = "0.00%"
                                Case "Chg % 1Y"
                                    ws1.Range(ws1.Cells(i, (j + 9)), ws1.Cells(i, (j + 9))) = objMyRecordset.Fields("Return_1Yr")
                                    ws1.Range(ws1.Cells(i, (j + 9)), ws1.Cells(i, (j + 9))).NumberFormat = "0.00%"
                                Case "Chg % 3Y"
                                    ws1.Range(ws1.Cells(i, (j + 9)), ws1.Cells(i, (j + 9))) = objMyRecordset.Fields("Return_3Yr")
                                    ws1.Range(ws1.Cells(i, (j + 9)), ws1.Cells(i, (j + 9))).NumberFormat = "0.00%"
                                Case "Chg % 3Y High"
                                    ws1.Range(ws1.Cells(i, (j + 9)), ws1.Cells(i, (j + 9))) = objMyRecordset.Fields("Return_3Yr_High")
                                    ws1.Range(ws1.Cells(i, (j + 9)), ws1.Cells(i, (j + 9))).NumberFormat = "0.00%"
                                Case "Div Yld"
                                    ws1.Range(ws1.Cells(i, (j + 9)), ws1.Cells(i, (j + 9))) = objMyRecordset.Fields("Exp_1Yr_Dividend_Yld")
                                    ws1.Range(ws1.Cells(i, (j + 9)), ws1.Cells(i, (j + 9))).NumberFormat = "0.00%"
                                Case "Div Yld Rank"
                                    ws1.Range(ws1.Cells(i, (j + 9)), ws1.Cells(i, (j + 9))) = objMyRecordset.Fields("Exp_1Yr_Dividend_Yld_Rank_Perc")
                                    ws1.Range(ws1.Cells(i, (j + 9)), ws1.Cells(i, (j + 9))).NumberFormat = "0"
                                Case Else
                                    k = 0
                            End Select

                    Next j
                    
                    i = i + 1
                    objMyRecordset.MoveNext
                Loop

            End If
    
            ' Close and clean up connections
            objMyRecordset.Close
            Set objMyRecordset = Nothing
            objMyConn.Close
            Set objMyConn = Nothing
            Exit Sub
    
        ExitHandler:

            ' Close and clean up connections
            objMyRecordset.Close
            Set objMyRecordset = Nothing
            objMyConn.Close
            Set objMyConn = Nothing
            Exit Sub

        ErrorHandler:
            MsgBox "Error during SQL execution: " & Err.Description
            Resume ExitHandler
    
        End Sub

        ' Sort the Tickers in B column
        Sub Sort_Tickers(start_row As Integer, end_row As Integer)
            Dim rng1 As Range
            Dim rng2 As Range
            Dim wb1 As Workbook
            Dim ws1 As Worksheet

            Set wb1 = ThisWorkbook
            Set ws1 = wb1.Worksheets("Equities")
            ws1.Activate

            ' Define the range that needs to be sorted which is column B
            Set rng1 = ws1.Range("C" & start_row & ":R" & end_row)
            ' Define the key range for sorting which is in column R
            Set rng2 = ws1.Range("R" & start_row & ":R" & end_row)

            ' Clear any previous sort fields
            ws1.Sort.SortFields.Clear

            ' Add the sorting key and sort
            ws1.Sort.SortFields.Add2 Key:=rng2, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            With ws1.Sort
                .SetRange rng1
                .Header = xlGuess
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With

        End Sub

**Continued...** [Part Three](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/readme3.md)



