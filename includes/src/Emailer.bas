Attribute VB_Name = "Emailer"

Sub TestEmailer()
    Emailer ("AW Carter LLC")
End Sub

Sub EmailAllTrueSubs()
    
    'START vLoop "v" TABLE LOOP (v is arbitrary for value, rename as desired)
    Dim vRow As Range
    Dim vRowIndex As Long
    Dim vTable As ListObject
    
    Set vTable = Worksheets("Emailer").ListObjects("Emailer_Sub_Status_Table")
    
    vRowIndex = 0
    For Each vRow In vTable.ListColumns("Sub").DataBodyRange.Rows
        vRowIndex = vRowIndex + 1
        'Use vRow as needed in the loop
        'vRow.Row
        If vTable.DataBodyRange.Cells(vRowIndex, vTable.ListColumns("Send Report").Index) = True Then
            e = Emailer(vTable.DataBodyRange.Cells(vRowIndex, vTable.ListColumns("Sub").Index), vTable.DataBodyRange.Cells(vRowIndex, vTable.ListColumns("Emails").Index))
        End If
    Next vRow
    ''''' END vLoop

End Sub


Function Emailer(subName As String, subEmails As String)
    
    'UPDATE DATA TO SUB
    Range("Filter_Sub_Name").Value = subName
    
    'GET ATTACHMENT
    subAttachmentFolder = Application.ActiveWorkbook.Path & "\includes\exports\" & Format(Range("Current_Data_Date").Value, "yyyy-MM-dd") & "\"
    attachmentFilePath = subAttachmentFolder & subName & " -- Status Update " & Format(Range("Current_Data_Date").Value, "yyyy-MM-dd") & ".pdf"
    MyMkDir (subAttachmentFolder)
    
    ThisWorkbook.Worksheets("Emailer_Sub_Data").ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        attachmentFilePath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        False

    'OPEN EMAIL
    e = SendEmail("DISPLAY", subName, subEmails, Worksheets("Emailer_Sub_Data").Range("G2").Value & "(" & Worksheets("Emailer_Sub_Data").Range("T6").Value & ")", Worksheets("Emailer_Sub_Data").Range("R2").Value, Worksheets("Emailer").Range("F3").Value, attachmentFilePath)
    
End Function



Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        '.Show
        On Error GoTo 0
    End With
    
    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function

Function DirExists(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    DirExists = GetAttr(DirName) And vbDirectory
ErrorHandler:
End Function

Function SendEmail(SENDorDISPLAYemail As String, subconName As Variant, subconEmail As Variant, emailSubject As String, emailBody As String, Optional ByVal ccEmail As String = "", Optional ByVal attachmentPath As String = "", Optional ByVal secondAttachmentPath As String = "")
    ' Requires Microsoft Outlook Object Library
    ' Based on http://www.exceltrainingvideos.com/automate-excel-to-pdf-and-email-pdf-document-using-vba/
    ' Camron 2019-03-12
    
    'Dim subconName As String
    'subconName = "Westland Construction"
    'Dim subconEmail As String
    'subconEmail = "camron@westlandconstruction.com"
    'Dim emailSubject As String
    'emailSubject = "Herriman High 2 Closeouts Test"
    'Dim emailBody As String
    'emailBody = "Please send me your closeouts ASAP"
    'Dim SENDorDISPLAYemail As String
    'SENDorDISPLAYemail = "SEND"
       
    Dim OutLookApp As Object
    Dim OutLookMailItem As Object
    Dim myAttachments As Object
    
    Set OutLookApp = CreateObject("Outlook.application")
    Set OutLookMailItem = OutLookApp.CreateItem(0)
    Set myAttachments = OutLookMailItem.Attachments
    
    Select Case SENDorDISPLAYemail
        Case Is = "DISPLAY"
            With OutLookMailItem
                .Importance = 2
                .To = subconEmail
                .CC = ccEmail
                .Subject = emailSubject
                .HTMLBody = emailBody
                
                If attachmentPath <> "" Then myAttachments.Add attachmentPath
                If secondAttachmentPath <> "" Then myAttachments.Add secondAttachmentPath
                .Display
            End With
        Case Is = "SEND"
            With OutLookMailItem
                .Importance = 2
                .To = subconEmail
                .CC = ccEmail
                .Subject = emailSubject
                .HTMLBody = emailBody
                If attachmentPath <> "" Then myAttachments.Add attachmentPath
                If secondAttachmentPath <> "" Then myAttachments.Add secondAttachmentPath
                .Send
            End With
        Case Else
        MsgBox ("SENDorDISPLAYemail is a required argument")
    End Select
    
    Set OutLookMailItem = Nothing
    Set OutLookApp = Nothing
End Function

Function GetBoiler(ByVal sFile As String) As String
    'Dick Kusleika https://www.rondebruin.nl/win/s1/outlook/signature.htm
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(sFile).OpenAsTextStream(1, -2)
    GetBoiler = ts.ReadAll
    ts.Close
End Function


