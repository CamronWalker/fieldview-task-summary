Attribute VB_Name = "Main"
Sub Resize_Task_Summary_Table()
    'Refresh Tasks Table
    Worksheets("Tasks_Import").ListObjects("Tasks").Refresh
    
    'Get resize length from imported Index
    wsLength = WorksheetFunction.Max(Sheets("Tasks_Import").ListObjects("Tasks").ListColumns("Index").Range)
        
    'Resize the table UPDATE COLUMNS IF NEEDED
    Sheets("Task_Summary").ListObjects("Task_Summary_Table").Resize Range("A1:I" & wsLength + 1)
End Sub
