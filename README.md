# Excel VBA Custom SQL Query Tool

This simple query tool was designed using Excel VBA macros, shapes, and Active X Controls. It queries an Azure SQL server database using encrypted username and password and query code stored in the cells of a hidden Excel tab called *Queries*.

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




