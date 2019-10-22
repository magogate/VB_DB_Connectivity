Option Explicit

Public cn As ADODB.Connection


Sub ProcessData()

'https://stackoverflow.com/questions/5349580/compiler-error-user-defined-types-not-defined
'https://stackoverflow.com/questions/18313899/vba-new-database-connection

Set cn = New ADODB.Connection
Dim rs As ADODB.Recordset
Dim StrSQL As String
Dim id As Integer
Dim Name As String
Dim code As String
Dim City As String
Dim wsName As String

Dim strConn As String
    strConn = "Driver={SQL Server};Server=MAGOGATE-PC; Database=AdventureWorks2017; UID=sa; PWD=magogate"

Dim lastRow As Double
Dim row As Double
Dim ws As Integer

    cn.Open strConn
    StrSQL = "Truncate table dbo.MyTable"
    cn.Execute StrSQL

    For ws = 1 To Worksheets.Count
        wsName = Worksheets(ws).Name
        'MsgBox (wsName)
        'activating each worksheet in order to iterate through data
         Worksheets(wsName).Activate
        
        'going to very first cell of worksheet to get last row to find out range
         Cells(1, 1).Select
    
        'fetching last row so that we can iterate through each & every cell
         lastRow = Cells(Rows.Count, 1).End(xlUp).row
         
         
         
         For row = 2 To lastRow
            
               id = Cells(row, 1).Value
               Name = Cells(row, 2).Value
               code = Cells(row, 3).Value
               City = Cells(row, 4).Value
            
               StrSQL = "insert into dbo.MyTable(id, name, code, city) values(" _
               + Str(id) + "," _
               + "'" + Name + "'," _
               + "'" + code + "'," _
               + "'" + City + "'" _
               + ")"
               
               'MsgBox (StrSQL)
               
               cn.Execute StrSQL
            
         Next row
         
         'MsgBox (StrSQL)
    
    
    Next
    cn.Close
    Set cn = Nothing
    MsgBox ("Process is finished...")

End Sub



