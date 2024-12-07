# Excel VBA Custom SQL Query Tool

This simple query tool was designed using Excel VBA macros, shapes, and Active X Controls. It queries an Azure SQL server database using encrypted username and password and query code stored in the cells of a hidden Excel sheet called *Queries*.

![Equity_Screening_Tool.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool.jpg?raw=true)

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

![Equity_Screening_Tool.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool.jpg?raw=true)

### 1. Date ComboBoxes

The following function called *Update_Dates* connects the database using the *Set_Conn* function. The SQL statements for Start Dates and End Dates are stored in cells in the *Queries* sheet. query_cell_1 stores the 1st cell reference for Start Dates and query_cell_2 stores the 2nd cell reference for End Dates. Basically, Date Arrays are used to store the dates that are then passed as Lists to the ComboBoxes. Finally, we cleanup and close the connections.

![Equity_Screening_Tool.jpg](https://github.com/danvuk567/Excel_VBA-Custom-SQL-Query-Tool/blob/main/images/Equity_Screening_Tool.jpg?raw=true)

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

