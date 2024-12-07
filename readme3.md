# Excel VBA Custom SQL Query Tool (Part Two)

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

The current query returns the results for the *Information Technology* sector and the *Semiconductors* Sub-Industry for December 6th 2024.

![Equity_Screening_Tool_Query_Results.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool_Query_Results.jpg?raw=true)
