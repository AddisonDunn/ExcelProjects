Sub CopyItem()
    Set ToWS = Worksheets("Products")
    Set originalCell = ActiveCell
    
    'If the first cell in the selection is over the products list
    If (12 < ActiveCell.Column) Then
        Dim objSelection As Range
        Dim objSelectionArea As Range
        Dim objCell As Range
        Dim intRow As Integer
        
        'get the selection that the user has made
        Set objSelection = Application.Selection
        ' Get the current selection (if there are mulitple--perhaps if the users selection skips lines)
        Set objSelection = Application.Selection

        ' Walk through the areas
        For Each objSelectionArea In objSelection.Areas

            ' Walk through the rows
            For intRow = 1 To objSelectionArea.Rows.Count Step 1
                'target last blank row in ordering sheet
                Set targetRange = Range("C" & (Range("C" & Rows.Count).End(xlUp).Row + 1))
                
                ' Get the row reference
                Set objCell = objSelectionArea.Rows(intRow)
                ' Get the actual row index (in the worksheet).
                ' The other row index is relative to the collection.
                intActualRow = objCell.Row
    
                
                'source one of the rows in the selection
                Set sourceRange = Range("N" & intActualRow)
                
                'copy product name and website from right side of order form to the left
                sourceRange.Copy (targetRange)
                targetRange.Offset(0, 1).Value = "0"
                sourceRange.Offset(0, 1).Copy (targetRange.Offset(0, 2))
                sourceRange.Offset(0, 6).Copy (targetRange.Offset(0, 7))
                
                'set the formulas for the data that was copied over (index/match lookup)
                targetRow = targetRange.Row
                targetRange.Offset(0, 3).Value = "=INDEX(Products!C2:G5000,MATCH(C" & targetRow & "&E" & targetRow & ", Products!A2:A5000, 0), 3)"
                targetRange.Offset(0, 4).Value = "=D" & targetRow & "*F" & targetRow
                targetRange.Offset(0, 5).Value = "=INDEX(Products!C2:G5000,MATCH(C" & targetRow & "&E" & targetRow & ", Products!A2:A5000, 0), 4)"
                targetRange.Offset(0, 6).Value = "=INDEX(Products!C2:G5000,MATCH(C" & targetRow & "&E" & targetRow & ", Products!A2:A5000, 0), 5)"
            
            Next
        Next
    End If
    
    'adds data validation for products and website to the left side of the order form
    
    With ToWS
        Set sourceRange = .Range("C" & (.Range("C" & .Rows.Count).End(xlUp).Row + 1))
    End With
    
    Set targetRange = Range("$C$2:C" & (Range("C" & Rows.Count).End(xlUp).Row + 1))
    With targetRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=("=Products!" & "$C$2:" & sourceRange.Address)
    End With
    
    Set targetRange2 = Range("$E$2:E" & (Range("C" & Rows.Count).End(xlUp).Row + 1))
    With targetRange2.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=("=Products!$N$8:$N$11")
    End With
    
    SortOrderForm
    
    original
