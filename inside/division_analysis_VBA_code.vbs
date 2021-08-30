'-------1.division_analysis  (topics or division analysis)--------
    Sub template_division_analysis()
        Dim Sheet1 As Worksheet
        Set Sheet1 = Sheets.Add
    
    '1.Dealing with columns
        Dim arrC As Variant
        Dim strc As String
        Dim i As Integer
        
        arrC = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
          
            'Dealing with individual columns
                ' "S.No." column
                Sheet1.Columns(arrC(i)).ColumnWidth = 5
                Sheet1.Columns(arrC(i)).WrapText = True
                strc = arrC(i) & CStr(2) & ":" & arrC(i) & CStr(7)
                Sheet1.Range(strc).Merge
                strc = arrC(i) & CStr(2)
                Sheet1.Range(strc).Value = "S.No."
                
                ' "Topics" column
                i = i + 1
                Sheet1.Columns(arrC(i)).ColumnWidth = 20
                Sheet1.Columns(arrC(i)).WrapText = True
                strc = arrC(i) & CStr(2) & ":" & arrC(i) & CStr(7)
                Sheet1.Range(strc).Merge
                strc = arrC(i) & CStr(2)
                Sheet1.Range(strc).Value = "Topics"
                
                ' "Main Sub-topics" column
                i = i + 1
                Sheet1.Columns(arrC(i)).ColumnWidth = 33
                Sheet1.Columns(arrC(i)).WrapText = True
                strc = arrC(i) & CStr(2) & ":" & arrC(i) & CStr(7)
                Sheet1.Range(strc).Merge
                strc = arrC(i) & CStr(2)
                Sheet1.Range(strc).Value = "Main Sub-topics"
                
                ' "Sub Sub-topics" column
                i = i + 1
                Sheet1.Columns(arrC(i)).ColumnWidth = 27
                Sheet1.Columns(arrC(i)).WrapText = True
                strc = arrC(i) & CStr(2) & ":" & arrC(i) & CStr(7)
                Sheet1.Range(strc).Merge
                strc = arrC(i) & CStr(2)
                Sheet1.Range(strc).Value = "Sub Sub-topics"
                
                ' "Other sub-topics" column
                i = i + 1
                Sheet1.Columns(arrC(i)).ColumnWidth = 40
                Sheet1.Columns(arrC(i)).WrapText = True
                strc = arrC(i) & CStr(2) & ":" & arrC(i) & CStr(7)
                Sheet1.Range(strc).Merge
                strc = arrC(i) & CStr(2)
                Sheet1.Range(strc).Value = "Other sub-topics"
                Sheet1.Columns(arrC(i)).Font.Color = RGB(255, 0, 0)
                strc = arrC(i) & CStr(1) & ":" & arrC(i) & CStr(9)
                Sheet1.Range(strc).Font.Color = RGB(0, 0, 0)
                
                ' "topics which u might miss" columns
                i = i + 1
                Sheet1.Columns(arrC(i)).ColumnWidth = 6.22
                Sheet1.Columns(arrC(i)).WrapText = True
                strc = arrC(i) & CStr(2) & ":" & arrC(i) & CStr(7)
                Sheet1.Range(strc).Merge
                strc = arrC(i) & CStr(2)
                Sheet1.Range(strc).Value = "topics which u might miss"
                
                ' "hard topics" columns
                i = i + 1
                Sheet1.Columns(arrC(i)).ColumnWidth = 5.33
                Sheet1.Columns(arrC(i)).WrapText = True
                strc = arrC(i) & CStr(2) & ":" & arrC(i) & CStr(7)
                Sheet1.Range(strc).Merge
                strc = arrC(i) & CStr(2)
                Sheet1.Range(strc).Value = "hard topics"
                
                ' "page numbers involved" column
                i = i + 1
                Sheet1.Columns(arrC(i)).ColumnWidth = 6
                Sheet1.Columns(arrC(i)).WrapText = True
                strc = arrC(i) & CStr(2) & ":" & arrC(i) & CStr(7)
                Sheet1.Range(strc).Merge
                strc = arrC(i) & CStr(2)
                Sheet1.Range(strc).Value = "page numbers involved"
                
                ' "total pages involved" column
                i = i + 1
                Sheet1.Columns(arrC(i)).ColumnWidth = 6
                Sheet1.Columns(arrC(i)).WrapText = True
                strc = arrC(i) & CStr(2) & ":" & arrC(i) & CStr(7)
                Sheet1.Range(strc).Merge
                strc = arrC(i) & CStr(2)
                Sheet1.Range(strc).Value = "total pages involved"
                
                ' BLANK COLUMN
                i = i + 1
                Sheet1.Columns(arrC(i)).ColumnWidth = 8.11
                
                ' "actual start page" column
                i = i + 1
                Sheet1.Columns(arrC(i)).ColumnWidth = 7.56
                Sheet1.Columns(arrC(i)).WrapText = True
                strc = arrC(i) & CStr(2) & ":" & arrC(i) & CStr(7)
                Sheet1.Range(strc).Merge
                strc = arrC(i) & CStr(2)
                Sheet1.Range(strc).Value = "actual start page"
                
                ' "actual end page" column
                i = i + 1
                Sheet1.Columns(arrC(i)).ColumnWidth = 7.56
                Sheet1.Columns(arrC(i)).WrapText = True
                strc = arrC(i) & CStr(2) & ":" & arrC(i) & CStr(7)
                Sheet1.Range(strc).Merge
                strc = arrC(i) & CStr(2)
                Sheet1.Range(strc).Value = "actual end page"
                
                ' "actual total pages" column
                i = i + 1
                Sheet1.Columns(arrC(i)).ColumnWidth = 9
                Sheet1.Columns(arrC(i)).WrapText = True
                strc = arrC(i) & CStr(2) & ":" & arrC(i) & CStr(7)
                Sheet1.Range(strc).Merge
                strc = arrC(i) & CStr(2)
                Sheet1.Range(strc).Value = "actual total pages"
                
                ' BLANK COLUMN
                i = i + 1
                Sheet1.Columns(arrC(i)).ColumnWidth = 8.11
                
                ' "difficulty level" column
                i = i + 1
                Sheet1.Columns(arrC(i)).ColumnWidth = 9
                Sheet1.Columns(arrC(i)).WrapText = True
                strc = arrC(i) & CStr(2) & ":" & arrC(i) & CStr(7)
                Sheet1.Range(strc).Merge
                strc = arrC(i) & CStr(2)
                Sheet1.Range(strc).Value = "difficulty level"
                
                ' "division of topics so that u can put in promise" column
                i = i + 1
                Sheet1.Columns(arrC(i)).ColumnWidth = 9
                Sheet1.Columns(arrC(i)).WrapText = True
                strc = arrC(i) & CStr(2) & ":" & arrC(i) & CStr(7)
                Sheet1.Range(strc).Merge
                strc = arrC(i) & CStr(2)
                Sheet1.Range(strc).Value = "division of topics so that u can put in promise"
                
                ' "max. time taken in that division" column
                i = i + 1
                Sheet1.Columns(arrC(i)).ColumnWidth = 9
                Sheet1.Columns(arrC(i)).WrapText = True
                strc = arrC(i) & CStr(2) & ":" & arrC(i) & CStr(7)
                Sheet1.Range(strc).Merge
                strc = arrC(i) & CStr(2)
                Sheet1.Range(strc).Value = "max. time taken in that division"
                
            'Dealing with all rows
                Sheet1.Columns("A:Q").HorizontalAlignment = xlCenter
                
                MsgBox ("hey")
                
    '2.Dealing with rows
            'row 1
                Sheet1.Range("A1:Q1").Merge
                
            'row 2
                Sheet1.Range("A2:Q2").Font.Bold = True
                
            'row 9
                Sheet1.Range("9:9").Interior.Color = RGB(0, 32, 96)
               
            'row 99
                Sheet1.Range("99:99").Interior.Color = RGB(255, 192, 0)
                
            'row 100
                Sheet1.Range("100:100").Interior.Color = RGB(0, 32, 96)
                
    '3.Dealing with cells
            'Cell A1
                Sheet1.Range("A1").Value = "Placeholder"
                Sheet1.Range("A1").Font.Bold = True
                    
    End Sub