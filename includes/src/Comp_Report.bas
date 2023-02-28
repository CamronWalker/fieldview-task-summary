Attribute VB_Name = "Comp_Report"
Private Sub GetCompletedBetweenDates()
    Dim conStartDate As Date
    Dim conFinishDate As Date
    

    conStartDate = Range("Past_Comparison_Data_Date").Value
    conStartSheetName = "TS_" & Format(conStartDate, "yyyy-MM-dd")
    conStartTableName = "TS_" & Format(conStartDate, "yyyyMMdd") & "_Table"
    
    conFinishDate = Range("Current_Data_Date").Value
    conFinishSheetName = "TS_" & Format(conFinishDate, "yyyy-MM-dd")
    conFinishTableName = "TS_" & Format(conFinishDate, "yyyyMMdd") & "_Table"
    
    ' For each row in finish table
    Dim fRow As Range 'for Row Value
    Dim fRowIndex As Long
    Dim fTable As ListObject
    
    Set fTable = Worksheets(conFinishSheetName).ListObjects(conFinishTableName)
    Set sTable = Worksheets(conStartSheetName).ListObjects(conStartTableName).ListColumns("Task Number").DataBodyRange
    
    fRowIndex = 0
    For Each fRow In fTable.ListColumns("Task Number").DataBodyRange.Rows
        fRowIndex = fRowIndex + 1
        ' Use fRow if you only need the value from that column
        ' Use comment below for different row values based on header name
        ' fTable.DataBodyRange.Cells(fRowIndex, fTable.ListColumns("Task Number").Index)
        
        Set findStart = sTable.Find(fRow.Value)
        
        If Not findStart Is Nothing Then 'prevents an error for nothing.
            If findStart.Row = fRow.Value Then
                'find the status row
                'if status in fTable is complete and on the sTable it isn't then push
            
            
            
            
            End If
        Else ' If it doesn't exist in the start we want the value pushed
            'push fRow.value
        
        End If
    Next fRow
    ''''''''END fLoop
    
End Sub
