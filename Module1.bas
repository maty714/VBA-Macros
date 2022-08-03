Attribute VB_Name = "Module1"
Option Explicit



Sub BOA_Classification()


'Declare/set variables
Dim Cll As Range
Dim lookupRange As Range
Dim entryStart As Integer
Dim i As Byte
Dim result_BH As Integer
Dim upperCase
entryStart = Range("E1").End(xlDown).Row

'Insert new column and set lookupRage
Range("H:H").Insert
Set lookupRange = Range("J2", Range("J" & Rows.Count).End(xlUp))



Dim A As total
Set A = New total

'Remove errors on amounts that were imported
A.ConvertError


'Classify each item based on origin
For Each Cll In lookupRange
    upperCase = UCase(Cll)
      
    If (InStr(upperCase, "HEALTH") > 0) Then
        Cll.Offset(, -2).Value = "BH"
    End If
Next


'Entry Creation

Range("E" & entryStart).Offset(4, 0).Value = "Regular Entry"
Range("E" & entryStart).Offset(6, 0).Value = "Bulletin Healthcare Receipts"
Range("E" & entryStart).Offset(7, 0).Value = "Bulletin Media Receipts"
Range("E" & entryStart).Offset(8, 0).Value = "Cision Receipts"
Range("E" & entryStart).Offset(9, 0).Value = "BI Commercial Subscription"
Range("E" & entryStart).Offset(5, 1).Value = "Checks"
Range("E" & entryStart).Offset(6, 1).Value = A.checkTotal
Range("E" & entryStart).Offset(5, 2).Value = "ACH"
Range("E" & entryStart).Offset(6, 2).Value = A.achTotal
Range("E" & entryStart).Offset(5, 3).Value = "Total"






End Sub
