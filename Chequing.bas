Attribute VB_Name = "Chequing"
Sub FormatChequing()
    Call FormatChequingbool(True)
End Sub
Sub FormatChequingwithoutDeletion()
    Call FormatChequingbool(False)
End Sub
Sub FormatChequingbool(deleteValues As Boolean)
'
' Formats info from chequing sheets
'

'   date
    Range("F13").Select
    ActiveCell.FormulaR1C1 = _
        "=DATE(LEFT(RC[-3],4),MID(RC[-3],5,2),RIGHT(RC[-3],2))"
    Selection.NumberFormat = "dd-mmm-yy"
        
'   name
    Range("G13").Select
    ActiveCell.FormulaR1C1 = "=RIGHT(RC[-2],LEN(RC[-2])-4)"
    Selection.NumberFormat = "General"
    
'   debits taken out of account
    Range("I13").Select
    ActiveCell.FormulaR1C1 = "=-(RC[-5]-ABS(RC[-5]))/2"
    
'   credits put into account
    Range("J13").Select
    ActiveCell.FormulaR1C1 = "=(RC[-6]+ABS(RC[-6]))/2"
    
'   "debit" to indicate which account these transactions were from
    Range("K13").Select
    ActiveCell.FormulaR1C1 = "debit"
    
'   Format money values
    Columns("I:J").Select
    Selection.NumberFormat = "_-$* #,##0.00_-;-$* #,##0.00_-;;_-@_-"

'   Select and fill
    Range("F13:K13").Select
    Call AutoFillVBA
    
'   delete credit card payments
    If deleteValues Then
        Call RemoveCreditCardPayments
    End If
End Sub
Sub RemoveCreditCardPayments()
    Dim numRows As Integer
    Dim j As Integer
    Dim i As Integer
    
    
'   delete transfers from debit card
    Range("G13").Select
    numRows = ActiveCell.End(xlDown).Row - 12   'number of transactions on credit card since last time
    
    For i = 1 To numRows
        j = numRows + 13 - i
        Cells(j, 7).Select
        If Cells(j, 7).Value = "redacted" Then
            Rows(j).EntireRow.Delete
        End If
    Next i
    
'   Select and Copy
    Range("F13:K13", Selection.End(xlDown)).Select
    Selection.Copy
End Sub




