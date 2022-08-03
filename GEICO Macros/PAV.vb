Option Explicit

Sub Automate_Email()

ScreenUpdating = False
Dim Filter_Range As Byte
Dim i As Byte
Dim Count_Row As Byte
Dim ArrRange As Range
Dim FCC_Range As Range
Dim LookupRange As Range
Dim Cll As Range, Eml As Range, Chk As Range
Dim myRange As Range
Dim Tbl As Range

Dim Yr As String, Dte As String, Mnth As String, Acct As String


Set LookupRange = ActiveWorkbook.Worksheets("Email List").Range("A2:E18")
Dim OutlookApp As Outlook.Application
Dim OutlookEmail As Outlook.MailItem


Dim PAV As Worksheet
Set PAV = Sheets("PAV")
i = 7

On Error GoTo EH
'Clears array list then sorts check number by FCC
Worksheets("Array List").Cells.Clear
Worksheets("PAV").Sort.SortFields.Clear
Range("B6").Sort Key1:=Range("B6"), Header:=xlYes


'input values create file path for attachments
Yr = InputBox("Select the yearly folder you want to access. Please enter a year:")
If Yr = "" Then Exit Sub

Mnth = InputBox("Please select a monthly folder you want to access. Please type a full months name:")
If Mnth = "" Then Exit Sub

Dte = InputBox("Please enter the fulldate for the folder you want to access. (I.e xx-xx-xxxx):")
If Dte = "" Then Exit Sub

With PAV
    Set FCC_Range = Worksheets("PAV").Range("B7", Range("B" & Rows.Count).End(xlUp))
    Filter_Range = Worksheets("PAV").Range("A6").End(xlDown).Row

    Do While Cells(i, 2).Value > 1

        .Cells(i, 4).Value = Application.WorksheetFunction.VLookup(Cells(i, 2), LookupRange, 4, False)
        i = i + 1

    Loop

End With



FCC_Range.Copy
Worksheets("Array List").Range("A1").PasteSpecial
ActiveWorkbook.Worksheets("Array List").Select

'First ArrRange is set to count all used cells and remove duplicates.
Set ArrRange = ActiveWorkbook.Worksheets("Array List").Range("A1", Range("A" & Rows.Count).End(xlUp))
ArrRange.RemoveDuplicates Columns:=1, Header:=xlNo

'ArrRange is set a second time to count the new cells used
Set ArrRange = ActiveWorkbook.Worksheets("Array List").Range("A1", Range("A" & Rows.Count).End(xlUp))
ActiveWorkbook.Worksheets("PAV").Select

'''''''''''''''''''''''''
'Start of Email Creation'
'''''''''''''''''''''''''
For Each Cll In ArrRange

    ActiveWorkbook.Worksheets("PAV").Range("B6").AutoFilter Field:=2, Criteria1:=Cll
    Set ChkRange = ActiveWorkbook.Worksheets("PAV").Range("A7", Range("A" & Rows.Count).End(xlUp))
    Set OutlookApp = New Outlook.Application
    Set OutlookEmail = OutlookApp.CreateItem(olMailItem)
    Count_Row = WorksheetFunction.CountA(Range("A6", Range("A6").End(xlDown))) + 6
    Set Tbl = Range(Cells(6, 1), Cells(Count_Row, 1))
    
    
    With OutlookEmail
        For Each Chk In ChkRange
        Filename = CStr(Chk.Value)
            If Chk.Offset(, 2).Value = "31283 CLAIMS" Then
               .Attachments.Add "Y:\Trecs Daily Claims Reports (31283)" & "\" & Yr & "\" & Mnth & "\" & Dte & "\" & "\" & "PAV" & "\" & CStr(Chk.Value) & ".pdf"
            ElseIf Chk.Offset(, 2).Value = "31280 CLAIMS" Then
                .Attachments.Add "Y:\Trecs Daily Claims Reports (31280)" & "\" & Yr & "\" & Mnth & "\" & Dte & "\" & "\" & "PAV" & "\" & CStr(Chk.Value) & ".pdf"
            End If
        Next Chk
        
        .BodyFormat = olFormatHTML
        '.Display
        
        ' RangetoHTML(tbl) is a function that is stored in Email_To_HTM
        .HTMLBody = "Good Morning," & "<br>" & "The attached checks were presented for payment, but are VOIDED in the system." & vbNewLine _
        & "Please confirm by 2pm" & "<br>" & RangetoHTML(Tbl) & "<br>" & "Please send confirmations to Amanda Carrigan and myself." & vbNewLine _
        & "<br>" & "If you have any questions, please let me know." & "<br>" & "Thanks" & .HTMLBody
        .To = Range("D7").Value
        
        .Subject = "FCC " & Range("B7").Value & "- Paid Against Void -" & Dte
        .send
    
        
    End With
    
    Rows(7 & ":" & Filter_Range).Select
    Selection.ClearContents
    Selection.Delete Shift:=xlUp
    
Next Cll

MsgBox "Success, emails have been sent!"

Exit Sub
EH:
MsgBox "Please ensure that the folders entered exist or are spelled correctly", , Err.Description

ScreenUpdating = True

End Sub

