# VBA-challenge

'Please see all complete files via this google drive link: https://drive.google.com/drive/folders/1J_OK_zprDv5Zim9GZLic35eiW3fcxdEo?usp=sharing

'SOURCES LISTED BELOW

'Code sourced and referenced based off of lesson plans done in class, as well as tutor session, which helped me pick typo error issues with my code.

'Code sourced for running multiple worksheets was found from online: extendoffice.com/documents/excel/5333-excel-run-macro-multiple-sheets.html
'the code was the following:

Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call RunCode
    Next
    Application.ScreenUpdating = True
End Sub
Sub RunCode()
    'your code here
End Sub



'Code was also sourced with the help of askBCS for the following to correct / debug my code.
    Range("K" & Summary_Table_Row).NumberFormat = "0.00%"


'Also with the help of askBCS to correlate the ticker and output for bonus question
 ''Find the Maximum/Decrease in the columns percentage change and total stock volume
    maxIncrease = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastrow_c1)), Range("K2:K" & lastrow_c1), 0)
    maxDecrease = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & lastrow_c1)), Range("K2:K" & lastrow_c1), 0)
    maxVolume = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & lastrow_c1)), Range("L2:L" & lastrow_c1), 0)

' Returning the correct ticker

    Range("P2") = Cells(maxIncrease + 1, 9)
    Range("P3") = Cells(maxDecrease + 1, 9)
    Range("P4") = Cells(maxVolume + 1, 9)

    


