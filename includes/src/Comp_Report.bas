Attribute VB_Name = "Comp_Report"
Private Sub GetCompletedBetweenDates()
    Dim conStartDate As Date
    Dim conFinishDate As Date
    Dim idCollection As Collection
    
    Set idCollection = New Collection
    
    urlStart = Range("Edit_URL").Value

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
            If fTable.DataBodyRange.Cells(fRowIndex, fTable.ListColumns("Status").Index) = "Signed Off" Then
                If Worksheets(conStartSheetName).Cells(findStart.Row, 2) <> "Signed Off" Then
                    'Debug.Print Worksheets(conStartSheetName).Cells(findStart.Row, 1) & " ||| Was: " & Worksheets(conStartSheetName).Cells(findStart.Row, 2) & "  ---  Now is: " & fTable.DataBodyRange.Cells(fRowIndex, fTable.ListColumns("Status").Index)
                    idCollection.Add fRow.Value
                    countCompleted = countCompleted + 1
                
                End If
            End If
        Else ' If it doesn't exist in the start we want the value pushed if it's already signed off
            'push fRow.value
            If fTable.DataBodyRange.Cells(fRowIndex, fTable.ListColumns("Status").Index) = "Signed Off" Then
                idCollection.Add fRow.Value
                countCompleted = countCompleted + 1
            End If
        End If
nextf:
    Next fRow
    ''''''''END fLoop
    
    Debug.Print countCompleted
    Worksheets("Comparison Report").Range("D6").Value = countCompleted 'ADJUST IF START ROW CHANGES
    
    Worksheets("Comparison Report").Rows(12 & ":" & Worksheets("Comparison Report").Rows.Count).Delete 'ADJUST IF START ROW CHANGES
    
    cCount = 10 'ADJUST IF START ROW CHANGES
    
    For Each c In idCollection
        cCount = cCount + 1
        Worksheets("Comparison Report").Range("A" & cCount).Value = c ' "=HYPERLINK(""" & urlStart & c & """, """ & c & """)"
    Next c
    
    'For Each Row in Table Column TEMPLATE
    'START vLoop "v" TABLE LOOP (v is arbitrary for value, rename as desired)
    Dim vRow As Range 'for Row Value
    Dim vRowIndex As Long
    Dim vTable As ListObject
    
    Set vTable = Worksheets("Comparison Report").ListObjects("Comparison_Report_Table")
    
    vRowIndex = 0
    For Each vRow In vTable.ListColumns("Task Number").DataBodyRange.Rows
        vRowIndex = vRowIndex + 1
        ' Use vRow if you only need the value from that column
        ' Use comment below for different row values based on header name
        
        vRow.Formula = "=HYPERLINK(""" & urlStart & vRow.Value & """, """ & vRow.Value & """)"
        vTable.DataBodyRange.Cells(vRowIndex, vTable.ListColumns("Status").Index).Formula = "=XLOOKUP([@[Task Number]], " & conFinishTableName & "[Task Number], " & conFinishTableName & "[Status])"
        vTable.DataBodyRange.Cells(vRowIndex, vTable.ListColumns("Due").Index).Formula = "=XLOOKUP([@[Task Number]], " & conFinishTableName & "[Task Number], " & conFinishTableName & "[Due])"
        vTable.DataBodyRange.Cells(vRowIndex, vTable.ListColumns("Task Type").Index).Formula = "=XLOOKUP([@[Task Number]], " & conFinishTableName & "[Task Number], " & conFinishTableName & "[Task Type])"
        vTable.DataBodyRange.Cells(vRowIndex, vTable.ListColumns("Description").Index).Formula = "=XLOOKUP([@[Task Number]], " & conFinishTableName & "[Task Number], " & conFinishTableName & "[Description])"
        vTable.DataBodyRange.Cells(vRowIndex, vTable.ListColumns("Building").Index).Formula = "=XLOOKUP([@[Task Number]], " & conFinishTableName & "[Task Number], " & conFinishTableName & "[Building])"
        vTable.DataBodyRange.Cells(vRowIndex, vTable.ListColumns("Level").Index).Formula = "=XLOOKUP([@[Task Number]], " & conFinishTableName & "[Task Number], " & conFinishTableName & "[Level])"
        vTable.DataBodyRange.Cells(vRowIndex, vTable.ListColumns("Area/Room").Index).Formula = "=XLOOKUP([@[Task Number]], " & conFinishTableName & "[Task Number], " & conFinishTableName & "[Area/Room])"
        vTable.DataBodyRange.Cells(vRowIndex, vTable.ListColumns("To Package").Index).Formula = "=XLOOKUP([@[Task Number]], " & conFinishTableName & "[Task Number], " & conFinishTableName & "[To Package])"
        vTable.DataBodyRange.Cells(vRowIndex, vTable.ListColumns("To Org").Index).Formula = "=XLOOKUP([@[Task Number]], " & conFinishTableName & "[Task Number], " & conFinishTableName & "[To Org])"
    
    
    Next vRow
    ''''''''END vLoop
End Sub
