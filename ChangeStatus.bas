Sub ChangeStatus()
    'keeps sheet from running methods that run whenever the a sheet is changed. This way we can edit
    'the sheet without disturbance from other methods trying to run at the same time.
    On Error GoTo ErrorHandler
    Application.EnableEvents = False

    Set originalCell = ActiveCell
    
    Dim status As String

    'status = name of the button that called this method
    status = Application.Caller
    
    'If the first cell in the selection is over the orders in progress list
    If (ActiveCell.Column < 10 And Not IsEmpty(ActiveCell)) Then
        Dim objSelectionArea As Range
        Dim objCell As Range
        Dim intRow As Integer
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

                'loction of the cell we want to change to the given status
                Set actionRange = Range("B" & intActualRow)
                actionRange.Value = status
                               
            Next
        Next
    End If
    
ErrorHandler:
    Application.EnableEvents = True
    
    MoveCompletedOrders
    
End Sub
