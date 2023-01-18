Attribute VB_Name = "Main"
Sub Resize_Task_Summary_Table()
    'Refresh Tasks Table
    Worksheets("Tasks_Import").ListObjects("Tasks").Refresh
    
    'Get resize length from imported Index
    wsLength = WorksheetFunction.Max(Sheets("Tasks_Import").ListObjects("Tasks").ListColumns("Index").Range)
        
    'Resize the table UPDATE COLUMNS IF NEEDED  <<----------------------------------------------------------------------  REMEMBER
    Sheets("Task_Summary").ListObjects("Task_Summary_Table").Resize Range("A1:J" & wsLength + 1)
End Sub


Sub Update_Date_Totals()
    Dim vRow As Range 'for Row Value
    Dim dateFoundCol, i As Long
    i = 3
    
    Annex_Opened = Range("B3").Value
    Annex_Completed = Range("B4").Value
    Annex_Signed_Off = Range("B5").Value
    
    Historic_Opened = Range("AH3").Value
    Historic_Completed = Range("AH4").Value
    Historic_Signed_Off = Range("AH5").Value
    
    Site_Opened = Range("H3").Value
    Site_Completed = Range("H4").Value
    Site_Signed_Off = Range("H5").Value
    
    West_Tower_Opened = Range("T3").Value
    West_Tower_Completed = Range("T4").Value
    West_Tower_Signed_Off = Range("T5").Value
    
    Other_Opened = Range("N3").Value
    Other_Completed = Range("N4").Value
    Other_Signed_Off = Range("N5").Value

    
    
    
    For Each vRow In Worksheets("Tables").ListObjects("Completed_Over_Time_Table").ListColumns("Date").DataBodyRange.Rows
        ' Use vRow as needed in the loop
        If vRow = Date Then
            dateFoundCol = i
        End If
        
        i = i + 1
    Next vRow
    
    If dateFoundCol Then
        
    Else
    
    End If
    
End Sub

Sub CopyToDatedTable()
    TurnOffFunctionality
    Dim TS_Table As ListObject
    Dim TS_DatedTable As ListObject
    Dim tableDateVar As Date
    Dim TS_TableName As String
    Dim dest As Range
    
    tableDateVar = Range("Table_Date").Value
    
    TS_TableName = "TS_" & Format(tableDateVar, "yyyyMMdd") & "_Table"
    TS_SheetName = "TS_" & Format(tableDateVar, "yyyy-MM-dd")
    
    Sheets.Add(, Sheets(Sheets.Count)).Name = TS_SheetName
    
    Set TS_Table = Worksheets("Task_Summary").ListObjects("Task_Summary_Table")
    
    Set dest = Worksheets(TS_SheetName).Range("A1")
    TS_Table.Range.Copy
    
    With dest
        .PasteSpecial xlPasteAll
        .PasteSpecial xlPasteColumnWidths
    End With
    
    dest.CurrentRegion.Copy
    dest.PasteSpecial xlPasteValues
    
    Set TS_DatedTable = Worksheets(TS_SheetName).ListObjects(1)
    TS_DatedTable.Name = TS_TableName
    
    Worksheets("Tables").ListObjects("TS_Table_Dates_Table").ListRows.Add.Range(1) = tableDateVar
    

    
    TurnOnFunctionality
End Sub


Sub SubConDataUpdate()
'TurnOffFunctionality
    'this exists because the filter fuction doesn't copy in hyperlinks :/ So I have to do this manually...
    ' =FILTER(INDIRECT("TS_"&TEXT(Current_Data_Date,"yyyy\_MM\_dd")&"_Table"), (INDIRECT("TS_"&TEXT(Current_Data_Date,"yyyy\_MM\_dd")&"_Table[To Org]")=$D$2)*(INDIRECT("TS_"&TEXT(Current_Data_Date,"yyyy\_MM\_dd")&"_Table[Status]")<>"Signed Off"))
        
    Dim vRow As Range 'for Row Value
    Dim dataDate As Date
    Dim urlStart As String
        
    dataDate = Range("Current_Data_date").Value
    dataDateTableName = "TS_" & Format(dataDate, "yyyyMMdd") & "_Table"
    dataDateSheetName = "TS_" & Format(dataDate, "yyyy-MM-dd")
    filterSubName = Range("Filter_Sub_Name").Value
    
    
    Worksheets("Emailer_Sub_Data").ListObjects("Emailer_Sub_Data_Item_List_Table").DataBodyRange.Rows.Delete
    
    'Worksheets(dataDateSheetName).ListObjects(dataDateTableName).Range.Select
    
    With Worksheets(dataDateSheetName).ListObjects(dataDateTableName).Range
        .AutoFilter Field:=10, Criteria1:=filterSubName
        .AutoFilter Field:=2, Criteria1:=Array("Completed", "Opened", "Past Due"), Operator:=xlFilterValues
    End With
    
    isEmptyList = errorCatchEmptyFilter(Worksheets(dataDateSheetName).ListObjects(dataDateTableName).DataBodyRange)
    
    Worksheets("Emailer_Sub_Data").ListObjects("Emailer_Sub_Data_Item_List_Table").ListRows.Add.Range(1) = "NO OPEN ITEMS FOUND"
    
    If isEmptyList = False Then
    
        'This find gets the address of the top left corner so that I can paste into it (this preserves the table)
        Set ndf = Worksheets("Emailer_Sub_Data").ListObjects("Emailer_Sub_Data_Item_List_Table").DataBodyRange.Find("NO OPEN ITEMS FOUND", LookIn:=xlValues)
        
        findAddress = ndf.Address
        
        Worksheets(dataDateSheetName).ListObjects(dataDateTableName).DataBodyRange.Copy
        
        Worksheets("Emailer_Sub_Data").Range(findAddress).PasteSpecial xlPasteValues
        
        Worksheets("Emailer_Sub_Data").Range("Emailer_Sub_Data_Item_List_Table[Status]:Emailer_Sub_Data_Item_List_Table[Due]").FormatConditions.Delete
        
        ' UPDATE THE H17 HERE IF THE COLUMN/ROW CHANGES........................................
        Worksheets("Emailer_Sub_Data").Range("Emailer_Sub_Data_Item_List_Table[Status]:Emailer_Sub_Data_Item_List_Table[Due]").FormatConditions.Add Type:=xlExpression, Formula1:="=$H17=""Past Due"""
        
        With Worksheets("Emailer_Sub_Data").Range("Emailer_Sub_Data_Item_List_Table[Status]:Emailer_Sub_Data_Item_List_Table[Due]").FormatConditions(1)
            .Font.Color = RGB(192, 0, 0)
            .Font.Bold = True
        End With
        
        'Fix Hyperlinks (I did this here because they weren't copying in from the other table correctly on some subs, this was a work around)
        urlStart = Range("Edit_URL").Value
        For Each vRow In Worksheets("Emailer_Sub_Data").ListObjects("Emailer_Sub_Data_Item_List_Table").ListColumns("Task Number").DataBodyRange.Rows
            vRow.Formula = "=HYPERLINK(""" & urlStart & vRow.Value & """, """ & vRow.Value & """)"
        Next vRow
        '''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    End If

    Range("Filter_Sub_Name").Select 'Only there because it's cleaner to select the filter cell after a new selection

    'UPDATE PRINT AREA
    With Worksheets("Emailer_Sub_Data").ListObjects("Emailer_Sub_Data_Item_List_Table")
        bottomRow = .Range.Rows(.Range.Rows.Count).Row
    End With
    
    Worksheets("Emailer_Sub_data").PageSetup.PrintArea = "$G$2:$P$" & bottomRow

'TurnOnFunctionality

End Sub


Function errorCatchEmptyFilter(ByRef rngstart As Range) As Boolean
' https://stackoverflow.com/questions/33259296/handle-no-cells-were-found-error-when-filtered-range-is-empty
    errorCatchEmptyFilter = False
    
    'here I get an error if there are no cells
        On Error GoTo hell
        Set rngFiltered = rngstart.SpecialCells(xlCellTypeVisible)
    
    Exit Function
    
hell:
    errorCatchEmptyFilter = True

End Function

