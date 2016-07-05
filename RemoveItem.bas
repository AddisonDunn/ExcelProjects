
Sub RemoveItem()
    Set originalCell = ActiveCell
    
    'If the first cell in the selection is over the products list
    If (ActiveCell.Column < 11) Then
        Dim objSelection As Range
        Dim objSelectionArea As Range
        Dim objCell As Range
        Dim intRow As Integer
        
        Set objSelection = Application.Selection
        ' Get the current selection
        Set objSelection = Application.Selection
    
        ' Walk through the areas
        For Each objSelectionArea In objSelection.Areas
    
            ' Walk through the rows
            For intRow = 1 To objSelectionArea.Rows.Count Step 1
                
                ' Get the row reference
                Set objCell = objSelectionArea.Rows(intRow)
                
                ' Get the actual row index (in the worksheet).
                intActualRow = objCell.Row

                Set sourceRange = Range(Range("C" & intActualRow), Range("J" & intActualRow))
                sourceRange.ClearContents
                               
            Next
        Next
    End If
    
    SortOrderForm
    
    If IsEmpty(Range("D2").Value) Then
        Range("D2").Activate
        ActiveCell.Value = "0"
        ActiveCell.Offset(0, 2).Value = "=INDEX(Products!C2:G5000,MATCH(C2&E2, Products!A2:A5000, 0), 3)"
        ActiveCell.Offset(0, 3).Value = "=D2*F2"
        ActiveCell.Offset(0, 4).Value = "=INDEX(Products!C2:G5000,MATCH(C2&E2, Products!A2:A5000, 0), 4)"
        ActiveCell.Offset(0, 5).Value = "=INDEX(Products!C2:G5000,MATCH(C2&E2, Products!A2:A5000, 0), 5)"
    End If
    
    originalCell.Activate
End Sub
