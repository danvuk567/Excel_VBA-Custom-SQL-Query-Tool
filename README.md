# Excel VBA Custom SQL Query Tool

![Equity_Screening_Tool.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool.jpg?raw=true)

This simple Excel query tool was designed using Excel VBA macros, shapes, and Active X Controls. It queries an Azure SQL server database using encrypted username and password and query code stored in the cells of a hidden Excel sheet called *Queries*. The data set references over 2000 stocks from Public Financial Markets and can be filtered by dates indicated by section #1, Sector. It can also be filtered by Sub-Industry and Equity Ticker using ComboBoxes indicted in section #2 or by typing in the Ticker in column "B" indictated in section #4. The Dynamic ComboBoxes indicated in section #3 allow for selection of the various measures in any order that can be included or excluded. The data can also be filtered by Market Cap Size where *Small Cap* represents < 2 Billion Market Cap, *Mid Cap* represents 2 Billion to 10 Billion Market Cap, and *Large Cap* represents > 10 Billion Market Cap.

## Encrypt and Decrypt Database Connection String Functions

The following function called *XorEncrypt* can be used to encrypt the data source server name, database name, username and password within a connection string in the following format: "Provider=MSOLEDBSQL; Data Source=MyServerName; Initial Catalog=MyDBName; User ID=SomeID; Password=SomePassword;". The connection string can be passed as the first parameter and the second parameter will be the key which we can store in cell "A1" of the *Queries* tab.

    Function XorEncrypt(ByVal sData As String, ByVal sKey As String) As String
      Dim l As Long 
      Dim i As Long 
      Dim byIn() As Byte 
      Dim SOut() As String 
      Dim byKey() As Byte
      
      If Len(sData) = 0 Or Len(sKey) = 0 Then 
        XorEncrypt = "Invalid argument(s) used": 
        Exit Function
      End If
      
      byIn = StrConv(sData, vbFromUnicode)
      
      ReDim SOut(LBound(byIn) To UBound(byIn))
      
      byKey = StrConv(sKey, vbFromUnicode)
      
      l = LBound(byKey)
      For i = LBound(byIn) To UBound(byIn) Step 1
         SOut(i) = byIn(i) Xor byKey(l)
         l = l + 1
         If l > UBound(byKey) Then 
           l = LBound(byKey)
          End If
      Next i
      XorEncrypt = Join(SOut, ",")
      
    End Function

We can use this to encrypt the connection string and store the key in cell "A1" and the encrypted connection string in cell "A2" of the *Queries* tab as follows.

    Dim sConn As String
    Dim wb1 As Workbook
    Dim ws1 As Worksheet
    Dim encrypt_key As String
    Dim encrypted_connection As String
    Dim sConn As String
 
    wb1 = ThisWorkbook
    ws1 = wb1.Sheets("Queries")
    ws1.Activate

    sConn = *"Provider=MSOLEDBSQL; Data Source=MyServerName; Initial Catalog=MyDBName; User ID=SomeID; Password=SomePassword;"*
    encrypted_connection = XorEncrypt(sConn, encrypt_key)
    ws1.Range("A1").Value = encrypt_key
    ws1.Range("A2").Value = encrypted_connection

Now, in order to decrypt the encrypted connection string, we can use the function called *XorDecrypt* which passes the encrypted string and the key.

    Function XorDecrypt(ByVal sData As String, ByVal sKey As String) As String
        Dim i As Long 
        Dim l As Long
        Dim byOut() As Byte
        Dim sIn() As String
        Dim byKey() As Byte
        
        If Len(sData) = 0 Or Len(sKey) = 0 Then 
          XorEncrypt = "Invalid argument(s) used": 
          Exit Function
        End If
        
        sIn = Split(sData, ",")
        
        ReDim byOut(LBound(sIn) To UBound(sIn))
        
        byKey = StrConv(sKey, vbFromUnicode)
        
        l = LBound(byKey)
        For i = LBound(sIn) To UBound(sIn) Step 1
           byOut(i) = Val(sIn(i)) Xor byKey(l)
           l = l + 1
           If l > UBound(byKey) Then 
             l = LBound(byKey)
           End If
        Next i
        XorDecrypt = StrConv(byOut, vbUnicode)
        
    End Function

Finally, we can create a public function *Set_Conn* in Module1 that we can call whenever using database connections in the queries.

    Public Function Set_Conn() As String
        Dim sConn As String
    
        sConn = XorDecrypt(ThisWorkbook.Sheets("Queries").Range("A2").Value, ThisWorkbook.Sheets("Queries").Range("A1").Value)
        Set_Conn = "Provider=MSOLEDBSQL;" & sConn & ";"
    
    End Function


## Equities Sheet Query Components and Data Filtering

In order to connect and query the database, we can use **Microsoft ActiveX Data Objects**. The library needs to be enabled in *Tools/References*.

![Equity_Screening_Tool_Library.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool_Library.jpg?raw=true)

### 1. Date ComboBoxes

The following procedure called *Update_Dates* connects the database using the *Set_Conn* function. The SQL statements for Start Dates and End Dates are stored in cells in the *Queries* sheet. query_cell_1 stores the 1st cell reference for Start Dates and query_cell_2 stores the 2nd cell reference for End Dates. Basically, Date arrays are used to store the dates that are then passed as Lists to the ComboBoxes. Finally, we cleanup and close the connections.

![Equity_Screening_Tool_Dates.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool_Dates.jpg?raw=true)

        ' Update the Date ComboBox with Dates
        Sub Update_Dates(query_cell_1 As String, query_cell_2 As String)
            Dim objMyConn As ADODB.Connection
            Dim objMyRecordset As ADODB.Recordset
            Dim stConn As String
            Dim strSQL As String
            Dim strSQL2 As String
            Dim wb1 As Workbook
            Dim ws1 As Worksheet
            Dim ws2 As Worksheet

            Set wb1 = ThisWorkbook
            Set ws1 = wb1.Worksheets("Equities")
            Set ws2 = wb1.Worksheets("Queries")
    
            On Error Resume Next
    
            DisableApplication
    
            ' Initialize connection objects
            Set objMyConn = New ADODB.Connection
            Set objMyRecordset = New ADODB.Recordset
            stConn = Set_Conn
            objMyConn.ConnectionString = stConn
    
            ' Fetch SQL statement from query_cell_1
            ws2.Activate
            strSQL = ws2.Range(query_cell_1).Value

            ' Get RecordSet for SQL Execution of strSQL
            With objMyConn
                .CursorLocation = adUseClient 'Necesary for creating disconnected recordset.
                .Open stConn 'Open connection.
                Set objMyRecordset = .Execute(strSQL)
            End With
    
            ' Initialize Date_Array based on number of records
            ReDim Date_Array(objMyRecordset.RecordCount - 1)
    
            ' Move throught the RecordSet of strSQL and populate the array called Date_Array
            objMyRecordset.MoveFirst
            i = 0
            While Not objMyRecordset.EOF
                Date_Array(i) = Format(objMyRecordset.Fields(0), "yyyy-mm-dd")
                i = i + 1
                objMyRecordset.MoveNext
            Wend
    
            ' Populate Start Date and End Date ComboBox List values from Date_Array
            ws1.Activate
            ActiveSheet.ComboBox_Start_Date.Clear
            ActiveSheet.ComboBox_Start_Date.List = Date_Array
            ActiveSheet.ComboBox_End_Date.Clear
            ActiveSheet.ComboBox_End_Date.List = Date_Array
    
            ' Get RecordSet for SQL Execution of strSQL2
            ws2.Activate
            strSQL2 = ws2.Range(query_cell_2).Value
            Set objMyRecordset = objMyConn.Execute(strSQL2)
    
            ' Set Start Date and End Date ComboBox initial values based on the first and only Recordset value of strSQL2
            ws1.Activate
            If Not objMyRecordset.EOF Then
                objMyRecordset.MoveFirst
                ActiveSheet.ComboBox_Start_Date.Value = Format(objMyRecordset.Fields(0), "yyyy-mm-dd")
                ActiveSheet.ComboBox_End_Date.Value = Format(objMyRecordset.Fields(0), "yyyy-mm-dd")
            End If
    
            ' Close and clean up connections
            objMyRecordset.Close
            Set objMyRecordset = Nothing
            objMyConn.Close
            Set objMyConn = Nothing
    
            EnableApplication
  
        End Sub

### 2. Sector, Sub-Industries and Equities ComboBoxes

![Equity_Screening_Tool_Sectors.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool_Sectors.jpg?raw=true)

The following procedures will load the Sector ComboBox, the Sub-industries ComboBox and the Equities ComboBox. The tables are related and there is a **one-to-many relationship** between *Sectors* and *Sub_industries*. Sub-Industries has a **one-to-many relationship** with *Equities*. The Sub-Industried ComboBox is updated based on what is chosen in the Sectors ComboBox and the Equities ComboBox is updated based on what is chosen in the Sub-Industries ComboBox.

        ' Update the Sectors ComboBox with Sectors
        Sub Update_Sectors(query_cell As String)
            Dim objMyConn As ADODB.Connection
            Dim objMyRecordset As ADODB.Recordset
            Dim stConn As String
            Dim strSQL As String
            Dim wb1 As Workbook
            Dim ws1 As Worksheet
            Dim ws2 As Worksheet
            Dim i As Integer
            Dim Sector_Array As Variant
    
            Set wb1 = ThisWorkbook
            Set ws1 = wb1.Worksheets("Equities")
            Set ws2 = wb1.Worksheets("Queries")
            ws1.Activate
    
            On Error Resume Next
    
            DisableApplication
    
            ' Initialize connection objects
            Set objMyConn = New ADODB.Connection
            Set objMyRecordset = New ADODB.Recordset
            stConn = Set_Conn
            objMyConn.ConnectionString = stConn
    
            ' Fetch SQL statement from query_cell
            ws2.Activate
            strSQL = ws2.Range(query_cell).Value
    
            ' Get RecordSet for SQL Execution of strSQL
            With objMyConn
                .CursorLocation = adUseClient 'Necesary for creating disconnected recordset.
                .Open stConn 'Open connection.
                Set objMyRecordset = .Execute(strSQL)
            End With

            ' Initialize array called Sector_Array based on number of records
            ReDim Sector_Array(objMyRecordset.RecordCount - 1, 0)

            ' Move throught the RecordSet of strSQL and populate Sector_Array
            objMyRecordset.MoveFirst
            i = 0
            While Not objMyRecordset.EOF
                Sector_Array(i, 0) = objMyRecordset.Fields(0)
                i = i + 1
                objMyRecordset.MoveNext
            Wend

            ' Populate Sector ComboBox List values from Sector_Array
            ws1.Activate
            ActiveSheet.ComboBox_Sectors.Clear
            ActiveSheet.ComboBox_Sectors.List = Sector_Array
    
            ' Set initial values for the Sector, Sub-Industries, and Equities ComboBoxes
            ActiveSheet.ComboBox_Sectors.Value = "All Sectors"
            ActiveSheet.ComboBox_Sub_Industries = "All Sub-Industries"
            ActiveSheet.ComboBox_Equities = "All Equities"

            ' Close and clean up connections
            objMyRecordset.Close
            Set objMyRecordset = Nothing
            objMyConn.Close
            Set objMyConn = Nothing
    
            EnableApplication
    
        End Sub

![Equity_Screening_Tool_Sub_Industries.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool_Sub_Industries.jpg?raw=true)

        ' Update the Sub-Industries ComboBox
        Sub Update_Sub_Industries(query_cell As String)
            Dim objMyConn As ADODB.Connection
            Dim objMyRecordset As ADODB.Recordset
            Dim stConn As String
            Dim strSQL As String
            Dim wb1 As Workbook
            Dim ws1 As Worksheet
            Dim ws2 As Worksheet
            Dim i As Integer
            Dim Sub_Industry_Array As Variant
    
            Set wb1 = ThisWorkbook
            Set ws1 = wb1.Worksheets("Equities")
            Set ws2 = wb1.Worksheets("Queries")

            On Error Resume Next
    
            ThisWorkbook.DisableApplication

            ' Initialize connection objects
            Set objMyConn = New ADODB.Connection
            Set objMyRecordset = New ADODB.Recordset
            stConn = Set_Conn
            objMyConn.ConnectionString = stConn
        
            ws1.Activate
            ' Reset Sub-Industries Combobox to "All Sub-Industries" if Sectors ComoboBox is "All Sectors"
            If ActiveSheet.ComboBox_Sectors.Value = "All Sectors" Then
    
                ReDim Sub_Industry_Array(0, 0)
                Sub_Industry_Array(0, 0) = "All Sub-Industries"
    
                ActiveSheet.ComboBox_Sub_Industries.Clear
                ActiveSheet.ComboBox_Sub_Industries.List = Sub_Industry_Array
        
            Else
                ' Set Sub-Industry query to include Sector condition
                strSQL = ws2.Range(query_cell).Value & "'" & ActiveSheet.ComboBox_Sectors.Value & "' ORDER BY 2,1"
        
                ' Get RecordSet for SQL Execution of strSQL
                With objMyConn
                    .CursorLocation = adUseClient 'Necessary for creating disconnected recordset.
                    .Open stConn 'Open connection.
                    Set objMyRecordset = .Execute(strSQL)
                End With
        
                ' Initialize array called Sub_Industry_Array based on number of records
                ReDim Sub_Industry_Array(objMyRecordset.RecordCount - 1, 0)

                ' Move throught the RecordSet of strSQL and populate Sub_Industry_Array
                objMyRecordset.MoveFirst
                i = 0
                While Not objMyRecordset.EOF
                    Sub_Industry_Array(i, 0) = objMyRecordset.Fields(0)
                    i = i + 1
                    objMyRecordset.MoveNext
                Wend

                ' Populate Sub-Industry ComboBox List values from Sub_Industry_Array
                ws1.Activate
                ActiveSheet.ComboBox_Sub_Industries.Clear
                ActiveSheet.ComboBox_Sub_Industries.List = Sub_Industry_Array

            End If
    
            ' Close and clean up connections
            objMyRecordset.Close
            Set objMyRecordset = Nothing
            objMyConn.Close
            Set objMyConn = Nothing
    
            ' Set initial value for Sub-Industries ComboBox
            ActiveSheet.ComboBox_Sub_Industries.Value = "All Sub-Industries"
    
            ws1.Range("A1").Select
    
            ThisWorkbook.EnableApplication
    
        End Sub

![Equity_Screening_Tool_Equities.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool_Equities.jpg?raw=true)

        ' Update the Equities ComboBox
        Sub Update_Equities(query_cell As String)
            Dim objMyConn As ADODB.Connection
            Dim objMyRecordset As ADODB.Recordset
            Dim stConn As String
            Dim strSQL As String
            Dim query_text As String
            Dim wb1 As Workbook
            Dim ws1 As Worksheet
            Dim ws2 As Worksheet
            Dim i As Integer
            Dim Equity_Array As Variant
    
            Set wb1 = ThisWorkbook
            Set ws1 = wb1.Worksheets("Equities")
            Set ws2 = wb1.Worksheets("Queries")
    
            On Error Resume Next
    
            ThisWorkbook.DisableApplication
    
            ' Initialize connection objects
            Set objMyConn = New ADODB.Connection
            Set objMyRecordset = New ADODB.Recordset
            stConn = Set_Conn
            objMyConn.ConnectionString = stConn

            ws1.Activate
            ' Reset Equity Combobox to "All Equities" if Sub-Industries ComoboBox is "All Sub-Industries"
            If ActiveSheet.ComboBox_Sub_Industries.Value = "All Sub-Industries" Then
    
                ReDim Equity_Array(0, 0)
                Equity_Array(0, 0) = "All Equities"
                ActiveSheet.ComboBox_Equities.Clear
                ActiveSheet.ComboBox_Equities.List = Equity_Array
        
            Else
                ' Set Equity query to include Sub-Industries condition
                strSQL = ws2.Range(query_cell).Value & "'" & ActiveSheet.ComboBox_Sub_Industries.Value & "'" & " ORDER BY 2,1"
        
                ' Get RecordSet for SQL Execution of strSQL
                With objMyConn
                    .CursorLocation = adUseClient 'Necessary for creating disconnected recordset.
                    .Open stConn 'Open connection.
                    Set objMyRecordset = .Execute(strSQL)
                End With

                ' Initialize array called Equity_Array based on number of records
                ReDim Equity_Array(objMyRecordset.RecordCount - 1, 0)

                ' Move throught the RecordSet of strSQL and populate Equity_Array
                objMyRecordset.MoveFirst
                i = 0
                While Not objMyRecordset.EOF
                    Equity_Array(i, 0) = objMyRecordset.Fields(0)
                    i = i + 1
                    objMyRecordset.MoveNext
                Wend

               ' Populate Equity ComboBox List values from Equity_Array
                ActiveSheet.ComboBox_Equities.Clear
                ActiveSheet.ComboBox_Equities.List = Equity_Array

            End If
    
            ' Close and clean up connections
            objMyRecordset.Close
            Set objMyRecordset = Nothing
            objMyConn.Close
            Set objMyConn = Nothing
    
            ' Set initial value for Equity ComboBox
            ActiveSheet.ComboBox_Equities.Value = "All Equities"
    
            ws1.Range("A1").Select
    
            ThisWorkbook.EnableApplication
    
        End Sub

### 3. Dynamic ComboBoxes

![Equity_Screening_Tool_Dynamic_Column_Selection.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool_Dynamic_Column_Selection.jpg?raw=true)

The dynamic ComboBoxes are initially populated by the *Update_Score_Fields_Equities* procedure. This procedure sets up an array called *Score_Fields_Array* with various measures. The measures *Chg % 1M*, *Chg % 3M*, *Chg % 6M*, *Chg % 1Y*, and *Chg % 3Y* represent the 1 Month % change in a stock price, 3 Month % change in a stock price, 6 Month % change in a stock price, 1 Year % change in a stock price, and 3 Year % change in a stock price. The measure *Chg % 3Y High* represents the highest overall change in the past 3 years. *Div Yld* represent the dividend yield (0 if none) and *Div Yld Rank* represent the how they rank amongst all stocks with highest dividend yield being ranked #1. The dynamic ComboBox lists are then populated by Score_Fields_Array.

        ' Update the dynamic ComboBoxes
        Sub Update_Score_Fields_Equities()
        Dim wb1 As Workbook
            Dim ws1 As Worksheet
            Dim Score_Fields_Array As Variant
    
            Set wb1 = ThisWorkbook
            Set ws1 = wb1.Worksheets("Equities")
            ws1.Activate

            ' Create an Array called Score_Fields_Array with 8 Options for the purpose of dynamic data querying
            ReDim Score_Fields_Array(9, 0)
            Score_Fields_Array(0, 0) = ""
            Score_Fields_Array(1, 0) = "Chg % 1M"
            Score_Fields_Array(2, 0) = "Chg % 3M"
            Score_Fields_Array(3, 0) = "Chg % 6M"
            Score_Fields_Array(4, 0) = "Chg % 1Y"
            Score_Fields_Array(5, 0) = "Chg % 3Y"
            Score_Fields_Array(6, 0) = "Chg % 3Y High"
            Score_Fields_Array(7, 0) = "Div Yld"
            Score_Fields_Array(8, 0) = "Div Yld Rank"

            ' Clear all dynamic data related ComboBoxes
            ActiveSheet.ComboBox_Score_Fields1.Clear
            ActiveSheet.ComboBox_Score_Fields2.Clear
            ActiveSheet.ComboBox_Score_Fields3.Clear
            ActiveSheet.ComboBox_Score_Fields4.Clear
            ActiveSheet.ComboBox_Score_Fields5.Clear
            ActiveSheet.ComboBox_Score_Fields6.Clear
            ActiveSheet.ComboBox_Score_Fields7.Clear
            ActiveSheet.ComboBox_Score_Fields8.Clear
 
            ' Populate all all dynamic data related ComboBoxes with Score_Fields_Array
            ActiveSheet.ComboBox_Score_Fields1.List = Score_Fields_Array
            ActiveSheet.ComboBox_Score_Fields2.List = Score_Fields_Array
            ActiveSheet.ComboBox_Score_Fields3.List = Score_Fields_Array
            ActiveSheet.ComboBox_Score_Fields4.List = Score_Fields_Array
            ActiveSheet.ComboBox_Score_Fields5.List = Score_Fields_Array
            ActiveSheet.ComboBox_Score_Fields6.List = Score_Fields_Array
            ActiveSheet.ComboBox_Score_Fields7.List = Score_Fields_Array
            ActiveSheet.ComboBox_Score_Fields8.List = Score_Fields_Array
    
        End Sub


