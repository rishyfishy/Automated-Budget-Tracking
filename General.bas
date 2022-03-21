Attribute VB_Name = "General"
Sub Increment()
' increments current cell value by 1

ActiveCell = ActiveCell.Value + 1

End Sub
Sub Decrement()
' Decrement current cell value by 1

ActiveCell = ActiveCell.Value - 1

End Sub
Sub AutoFillVBA()
'
' Macro1 Macro
'

'
    Dim cell As Range
    Dim rng As Range
    
    Dim isFirstColumn As Boolean
    Dim isLastColumn As Boolean
    Dim isLastRow As Boolean
    
    Dim lastRow As Long
    Dim lastColumn As Long
    
    Dim leftEmpty As Boolean    'cells left of selection are empty
    Dim rightEmpty As Boolean   'cells right of selection are empty
    Dim downEmpty As Boolean    'cells immediately below selection are empty
    Dim downNotFull As Boolean  'cells immediately below selection are not ALL full,
    Dim leftBigger As Boolean   'column left of selection has more rows than column to the right
       
    isFirstColumn = Selection.Column = 1
    isLastColumn = Selection.Cells(1, Selection.Columns.Count).Column = Columns.Count
    isLastRow = Selection.Cells(Selection.Rows.Count, 1) = Rows.Count
    
    If isLastRow Then
        Exit Sub
    End If

    If isEmptyRange(Selection) Then Exit Sub
    
    
    downEmpty = isEmptyRange(Selection.Rows _
                (Selection.Rows.Count).Offset(1, 0))
                
    lastRow = Rows.Count + 1
    If downEmpty Then
    
        For Each cell In Selection.Rows(Selection.Rows.Count).Cells
            If Not IsEmpty(cell.End(xlDown)) Then
                lastRow = min(lastRow, cell.End(xlDown).Row - 1)
            End If
        Next cell
        If lastRow = Rows.Count + 1 Then
            GoTo LeftorRight
        Else
            GoTo Fill
        End If
    End If
    
    For Each cell In Selection.Rows(Selection.Rows.Count).Cells
        If downNotFull Then
            Exit For
        End If
        
        If IsEmpty(cell.Cells(2, 1)) Then
            downNotFull = True
        Else
            lastRow = min(lastRow, cell.End(xlDown).Row)
        End If
    Next cell
    
    If Not downNotFull Then

        GoTo Fill
    End If
    
    'By here, bottom should be empty
    
LeftorRight:
    If Not isFirstColumn Then
        leftEmpty = IsEmpty(Selection.Cells(1, 0)) And _
                    IsEmpty(Selection.Cells(2, 0))

        If Not leftEmpty Then
            lastRow = Selection.Cells(1, 0).End(xlDown).Row
        End If
    End If
    
    If Not isLastColumn Then
        rightEmpty = IsEmpty(Selection.Cells(1, Selection.Columns.Count + 1)) And _
                     IsEmpty(Selection.Cells(2, Selection.Columns.Count + 1))
        
        If Not rightEmpty Then
            lastRow = max(lastRow, Selection.Cells(1, Selection.Columns.Count + 1).End(xlDown).Row)
        End If
    End If
    
    If leftEmpty And rightEmpty Then Exit Sub

Fill:
        
    Set rng = Range(Selection.Cells(1, 1), _
        Cells(lastRow, Selection.Cells(1, Selection.Columns.Count).Column))
    Selection.AutoFill Destination:=rng
    rng.Select
    Exit Sub
End Sub
Public Function max(x, y As Variant) As Variant
  max = IIf(x > y, x, y)
End Function
Public Function min(x, y As Variant) As Variant
   min = IIf(x < y, x, y)
End Function
Public Function isEmptyRange(rng As Range) As Boolean
    Dim cell As Range
    isEmptyRange = True
    For Each cell In rng.Cells
        If Not IsEmpty(cell) Then
            isEmptyRange = False
            Exit Function
        End If
    Next cell
        
End Function






