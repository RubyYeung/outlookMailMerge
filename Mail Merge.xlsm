Sub Step1Clear()

Dim slSheet As Worksheet

'Set s1Sheet = Workbooks("StockScreen.xlsm").Sheets("TimeStampWork")
Set s1Sheet = ThisWorkbook.Sheets("Multi Group ID")
With s1Sheet
    'Range("A2:K" & slSheet.Cells(s1Sheet.Rows.Count, "A").End(xlUp).Row).ClearContents
     'Range("A2:K" & slSheet.Cells(s1Sheet.Rows.Count, "A").Row).ClearContents
   Range("A2:K" & s1Sheet.Cells(s1Sheet.Rows.Count, "A").Row).ClearContents
End With
End Sub

' This is an Excel macro
' you need to set a reference to Outlook Object Library
Public Sub Step2MultipleLoginMailMergeExcel()

 Dim olApp As Object
 Dim xlApp As Excel.Application
 Dim xlWB As Excel.Workbook
 Dim xlSheet As Excel.Worksheet
 Dim rCount As Long
 Dim bXStarted As Boolean
 Dim enviro As String
 Dim appdata As String
 Dim strPath As String
 Dim strAttachPath As String
 Dim SendTo As String
 Dim CCTo As String
 Dim strSubject As String
 Dim strAcctMgrName As String
 Dim AcctMgrEmail
 
 Dim olItem As Outlook.MailItem
 Dim Recip As Outlook.Recipient
 
' Get Excel set up
enviro = CStr(Environ("USERPROFILE"))
appdata = CStr(Environ("appdata"))
     On Error Resume Next
     Set xlApp = Excel.Application
     On Error GoTo 0
     'Open the workbook to input the data
     'Set xlWB = xlApp.ActiveWorkbook
     Set xlWB = xlApp.ThisWorkbook
     Set xlSheet = xlWB.Sheets("Multi Group ID")
    ' Process the message record
    
    On Error Resume Next

rCount = 2
strAttachPath = enviro & "\Documents\Send\"
'strAttachPath = Application.ThisWorkbook.Path & "\Documents\Send\"
Set olApp = GetObject(, "Outlook.Application")
     If Err <> 0 Then
         Set olApp = CreateObject("Outlook.Application")
         bXStarted = True
     End If

Do Until Trim(xlSheet.Range("A" & rCount)) = ""

strGroupID = xlSheet.Range("A" & rCount)
SendTo1 = xlSheet.Range("J" & rCount)
SendTo2 = xlSheet.Range("K" & rCount)
strSubject = xlSheet.Range("M" & 2)
' if adding attachment
'strAttachment = strAttachPath & xlSheet.Range("E" & rCount)
strLoginDate = xlSheet.Range("E" & rCount)
SenderEmail = xlSheet.Range("L" & 2)

'Create Mail Item and view before sending
' Default message form
'Set olItem = olApp.CreateItem(olMailItem)

' use a Template
'Set olItem = olApp.CreateItemFromTemplate(appdata & "\Microsoft\Templates\macro-test.oft")

Set olItem = olApp.CreateItemFromTemplate(Application.ThisWorkbook.Path & "\Multi group ID mail marge templateV2.oft")
    

With olItem

.SentOnBehalfOfName = SenderEmail
.To = SendTo1 & ";" & SendTo2
'.CC = CCTo
.Subject = strSubject


'.Body = Replace(.Body, "[GroupID]", strGroupID)
'.Body = Replace(.Body, "[LoginDate]", strLoginDate)

'.BodyFormat = olFormatHTML
'.HTMLBody = Replace(.HTMLBody, "[GroupID]", strGroupID)
'.HTMLBody = Replace(.HTMLBody, "[LoginDate]", strLoginDate)

.HTMLBody = Replace(.HTMLBody, "GroupID1", strGroupID)
.HTMLBody = Replace(.HTMLBody, "LoginDate1", strLoginDate)

'.Body = Replace(.Body, "[strAcctMgrName]", strAcctMgrName)
'if adding attachments:
'.Attachments.Add strAttachment
.Save

.Display
'.Send
End With

  rCount = rCount + 1
  
  Loop

Set xlWB = Nothing
Set xlApp = Nothing
     
End Sub
