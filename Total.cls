VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Total"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public total As Currency
Dim Cll As Range
Public amountRange As Range


Function achTotal()
    Set amountRange = Range("G2", Range("G" & Rows.Count).End(xlUp))
    total = 0

    For Each Cll In amountRange
        If (Cll.Offset(, 1).Value = "BH") Then
            total = total + Cll.Value
        End If
    Next Cll
    
    achTotal = total

End Function


Function checkTotal()
    Set amountRange = Range("G2", Range("G" & Rows.Count).End(xlUp))
    total = 0
    
    For Each Cll In amountRange
        If (Cll.Offset(, -1).Value = "Cash Letter Pre-Encoded Dep CR") Then
            total = total + Cll.Value
        End If
    Next Cll
    
    checkTotal = total

End Function


Sub ConvertError()
    Set amountRange = Range("G2", Range("G" & Rows.Count).End(xlUp))

    For Each Cll In amountRange
        Cll.Value = Evaluate(Cll.Address & "*1")
    Next

End Sub
