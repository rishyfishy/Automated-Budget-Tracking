Attribute VB_Name = "CreditCard"
Sub FormatCreditCard()

'
' Formats info from chequing sheets
'

'
    Dim numRows As Integer
    Dim j As Integer
    Dim i As Integer
    
    
'   delete transfers from debit card
    Range("A2").Select
    numRows = ActiveCell.End(xlDown).Row - 1
    For i = 1 To numRows
        j = numRows + 2 - i
        If Cells(j, 3).Value = "redacted" Then
            Rows(j).EntireRow.Delete
        End If
'   format CR values
        If Right(Cells(j, 4).Value, 2) = "CR" Then
            Cells(j, 4).Value = -1 * CDbl(Mid(Cells(j, 4), 2, Len(Cells(j, 4).Value) - 3))
        End If
    Next i
    
'   posting date
    Range("E2").Select
    ActiveCell.FormulaR1C1 = _
        "=DATE(RIGHT(RC[-3],4),MID(RC[-3],4,2),LEFT(RC[-3],2))+0.5"
    Selection.NumberFormat = "dd-mmm-yy"
'   name
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]"
    Selection.NumberFormat = "General"
    
'   Money taken out of credit card
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "=(RC[-4]+ABS(RC[-4]))/2"
    
'   Money paid to credit card
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=-(RC[-5]-ABS(RC[-5]))/2"
    
'   Format money values
    Columns("H:I").Select
    Selection.NumberFormat = "_-$* #,##0.00_-;-$* #,##0.00_-;;_-@_-"
    
'   "credit" to indicate which account these transactions were from
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "credit"
    Selection.NumberFormat = "General"
    
'   Select
    Range("E2:J2").Select
    Call AutoFillVBA
    
    

End Sub









